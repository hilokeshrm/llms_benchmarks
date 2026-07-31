"""
Benchmarks the models in dgx_llm_benchmark.xlsx through Ollama and appends the
results to Perf_Results.

    python run_ollama_bench.py --check                      # resolve tags, pull nothing
    python run_ollama_bench.py --priority P0 --pull         # pull + benchmark the P0 set
    python run_ollama_bench.py --models qwen3-32b,phi-4 \
        --concurrency 1,4,8 --shapes 2048/256 --repeats 3
    python run_ollama_bench.py --import-only results/perf_raw.jsonl

WHAT IT MEASURES
  Each "load point" is one (model x input/output shape x concurrency) combination,
  repeated --repeats times. A load point drives `concurrency` async workers in a
  closed loop against POST /api/generate with stream=true, timing every token
  arrival, while a background thread samples nvidia-smi for power and VRAM.
  One workbook row = one repeat of one load point.

CAVEAT
  Ollama is llama.cpp: GGUF weights, layers split across GPUs (pipeline), no
  tensor parallelism. Throughput here is Ollama's ceiling, not the DGX's. Every
  row is stamped serving_stack=Ollama so it never gets compared to a vLLM row.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import platform
import random
import statistics
import subprocess
import sys
import threading
import time
from datetime import datetime, timezone
from pathlib import Path

import httpx

import sheet_io

WORKBOOK = "dgx_llm_benchmark.xlsx"
RESULTS_DIR = Path("results")
NS = 1_000_000_000

# Filler corpus for synthetic prompts - neutral prose, no instructions the model
# might try to follow, which would change output length and pollute the timings.
CORPUS = (
    "the system records each measurement in a table before the next cycle begins "
    "operators review the values and compare them against the previous run while "
    "the schedule advances through every configuration in a fixed order without "
    "interruption so that later analysis can reconstruct exactly what happened at "
    "any point during the sequence of observations collected by the instruments "
).split()


# --------------------------------------------------------------- helpers ---
def utc_now():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def pct(values, p):
    """Percentile with linear interpolation. Returns None for an empty list."""
    if not values:
        return None
    xs = sorted(values)
    if len(xs) == 1:
        return xs[0]
    k = (len(xs) - 1) * (p / 100.0)
    lo, hi = int(k), min(int(k) + 1, len(xs) - 1)
    return xs[lo] + (xs[hi] - xs[lo]) * (k - lo)


def r2(x, n=2):
    return None if x is None else round(x, n)


def parse_list(s, cast=str):
    if s is None:
        return []
    if not isinstance(s, str):
        s = str(s)
    return [cast(x.strip()) for x in s.replace(";", ",").split(",") if x.strip()]


def parse_shapes(s):
    """'128/128, 2048/256' -> [(128,128), (2048,256)]"""
    out = []
    for item in parse_list(s):
        if "/" not in item:
            continue
        a, b = item.split("/", 1)
        out.append((int(a.strip()), int(b.strip())))
    return out


# ----------------------------------------------------------- GPU sampling ---
class GpuSampler(threading.Thread):
    """Polls nvidia-smi in the background. Degrades to a no-op if it is absent."""

    QUERY = "index,name,utilization.gpu,memory.used,power.draw"

    def __init__(self, interval=1.0):
        super().__init__(daemon=True)
        self.interval = interval
        self._stop = threading.Event()
        self.samples = []          # [(t, {idx: (util, mem_mb, watts)})]
        self.names = {}
        self.available = self._probe() is not None

    @classmethod
    def _probe(cls):
        try:
            out = subprocess.run(
                ["nvidia-smi", f"--query-gpu={cls.QUERY}", "--format=csv,noheader,nounits"],
                capture_output=True, text=True, timeout=15,
            )
            return out.stdout if out.returncode == 0 else None
        except (OSError, subprocess.SubprocessError):
            return None

    def _sample(self):
        raw = self._probe()
        if raw is None:
            return None
        snap = {}
        for line in raw.strip().splitlines():
            parts = [p.strip() for p in line.split(",")]
            if len(parts) < 5:
                continue
            try:
                idx = int(parts[0])
                self.names[idx] = parts[1]
                snap[idx] = (float(parts[2]), float(parts[3]), float(parts[4]))
            except ValueError:
                continue
        return snap

    def baseline(self):
        return self._sample() or {}

    def run(self):
        while not self._stop.is_set():
            snap = self._sample()
            if snap:
                self.samples.append((time.time(), snap))
            self._stop.wait(self.interval)

    def stop(self):
        self._stop.set()
        self.join(timeout=5)

    def aggregate(self, baseline=None):
        """Summarise the window. 'Active' GPUs are those this run actually loaded."""
        empty = dict(gpu_type=None, gpu_count=None, gpu_mem_peak_gb=None,
                     gpu_util_avg_pct=None, gpu_power_avg_w=None, gpu_power_peak_w=None)
        if not self.samples:
            return empty
        baseline = baseline or {}
        idxs = sorted({i for _, snap in self.samples for i in snap})
        peak_mem, max_util = {}, {}
        for i in idxs:
            mems = [snap[i][1] for _, snap in self.samples if i in snap]
            utils = [snap[i][0] for _, snap in self.samples if i in snap]
            peak_mem[i] = max(mems) if mems else 0.0
            max_util[i] = max(utils) if utils else 0.0
        active = [
            i for i in idxs
            if (peak_mem[i] - baseline.get(i, (0, 0, 0))[1]) > 512 or max_util[i] > 10
        ]
        if not active:
            active = idxs
        utils, powers = [], []
        for _, snap in self.samples:
            present = [i for i in active if i in snap]
            if not present:
                continue
            utils.append(statistics.fmean(snap[i][0] for i in present))
            powers.append(sum(snap[i][2] for i in present))
        return dict(
            gpu_type=self.names.get(active[0]),
            gpu_count=len(active),
            gpu_mem_peak_gb=r2(max(peak_mem[i] for i in active) / 1024.0),
            gpu_util_avg_pct=r2(statistics.fmean(utils)) if utils else None,
            gpu_power_avg_w=r2(statistics.fmean(powers)) if powers else None,
            gpu_power_peak_w=r2(max(powers)) if powers else None,
        )


# ---------------------------------------------------------- Ollama client ---
class Ollama:
    def __init__(self, host, timeout=1800.0):
        self.host = host.rstrip("/")
        self.timeout = timeout

    # -- sync control-plane calls -------------------------------------------
    def _get(self, path):
        with httpx.Client(timeout=60.0) as c:
            r = c.get(f"{self.host}{path}")
            r.raise_for_status()
            return r.json()

    def _post(self, path, payload, timeout=120.0):
        with httpx.Client(timeout=timeout) as c:
            r = c.post(f"{self.host}{path}", json=payload)
            r.raise_for_status()
            return r.json()

    def alive(self):
        try:
            self._get("/api/tags")
            return True
        except Exception:
            return False

    def local_tags(self):
        try:
            return {m["name"]: m for m in self._get("/api/tags").get("models", [])}
        except Exception:
            return {}

    def show(self, tag):
        try:
            return self._post("/api/show", {"model": tag})
        except Exception:
            return {}

    def unload(self, tag):
        """keep_alive=0 evicts the model so the next run measures a clean load."""
        try:
            self._post("/api/generate", {"model": tag, "prompt": "", "keep_alive": 0}, timeout=120.0)
        except Exception:
            pass

    def pull(self, tag, quiet=False):
        """Streams /api/pull. Returns (ok, seconds, error)."""
        t0 = time.perf_counter()
        last = 0.0
        try:
            with httpx.Client(timeout=None) as c:
                with c.stream("POST", f"{self.host}/api/pull",
                              json={"model": tag, "stream": True}) as resp:
                    resp.raise_for_status()
                    for line in resp.iter_lines():
                        if not line:
                            continue
                        obj = json.loads(line)
                        if "error" in obj:
                            return False, time.perf_counter() - t0, obj["error"]
                        if not quiet and obj.get("total"):
                            done = obj.get("completed", 0)
                            frac = done / obj["total"] * 100
                            now = time.perf_counter()
                            if now - last > 5:
                                last = now
                                print(f"      pulling {tag}: {frac:5.1f}%", end="\r", flush=True)
            if not quiet:
                print(" " * 60, end="\r")
            return True, time.perf_counter() - t0, None
        except Exception as e:
            return False, time.perf_counter() - t0, f"{type(e).__name__}: {e}"

    def registry_status(self, tag):
        """Best-effort remote existence check without downloading anything."""
        if tag.startswith("hf.co/") or tag.startswith("huggingface.co/"):
            return "unknown-hf-gguf"
        name, _, ver = tag.partition(":")
        ver = ver or "latest"
        repo = name if "/" in name else f"library/{name}"
        try:
            with httpx.Client(timeout=30.0, follow_redirects=True) as c:
                r = c.get(f"https://registry.ollama.ai/v2/{repo}/manifests/{ver}")
            if r.status_code == 200:
                return "available"
            if r.status_code in (401, 404):
                return "not-found"
            return f"http-{r.status_code}"
        except Exception:
            return "unknown"

    # -- the measured request ------------------------------------------------
    async def generate(self, client, tag, prompt, options, keep_alive="30m"):
        payload = {"model": tag, "prompt": prompt, "stream": True,
                   "options": options, "keep_alive": keep_alive}
        rec = {"ok": False, "ttft_s": None, "e2e_s": None, "in_tokens": 0,
               "out_tokens": 0, "itls": [], "load_s": None, "error": None}
        t0 = time.perf_counter()
        prev = None
        try:
            async with client.stream("POST", f"{self.host}/api/generate", json=payload) as resp:
                if resp.status_code != 200:
                    body = (await resp.aread()).decode("utf-8", "replace")[:200]
                    rec["error"] = f"HTTP {resp.status_code}: {body}"
                    return rec
                async for line in resp.aiter_lines():
                    if not line.strip():
                        continue
                    try:
                        obj = json.loads(line)
                    except json.JSONDecodeError:
                        continue
                    if obj.get("error"):
                        rec["error"] = str(obj["error"])[:300]
                        return rec
                    now = time.perf_counter()
                    if obj.get("response"):
                        if rec["ttft_s"] is None:
                            rec["ttft_s"] = now - t0
                        elif prev is not None:
                            rec["itls"].append(now - prev)
                        prev = now
                    if obj.get("done"):
                        rec["e2e_s"] = now - t0
                        rec["in_tokens"] = obj.get("prompt_eval_count") or 0
                        rec["out_tokens"] = obj.get("eval_count") or 0
                        if obj.get("load_duration"):
                            rec["load_s"] = obj["load_duration"] / NS
                        rec["ok"] = rec["out_tokens"] > 0
        except Exception as e:
            rec["error"] = f"{type(e).__name__}: {e}"
        if rec["e2e_s"] is None:
            rec["e2e_s"] = time.perf_counter() - t0
        return rec


# ------------------------------------------------------ prompt calibration ---
def build_prompt(n_words, seed=0):
    rng = random.Random(seed)
    return " ".join(rng.choice(CORPUS) for _ in range(n_words))


async def calibrate_prompt(ol, client, tag, target_isl, num_ctx, tol=0.03, max_iter=4):
    """Ollama gives no tokenizer endpoint, so we converge on the target prompt
    length using prompt_eval_count reported by a 1-token generation."""
    n_words = max(4, int(target_isl * 0.72))
    prompt, actual = build_prompt(n_words), None
    for _ in range(max_iter):
        prompt = build_prompt(n_words)
        rec = await ol.generate(client, tag, prompt,
                                {"num_predict": 1, "temperature": 0, "num_ctx": num_ctx})
        actual = rec["in_tokens"]
        if not actual:
            return prompt, None
        if abs(actual - target_isl) / target_isl <= tol:
            break
        n_words = max(4, int(n_words * target_isl / actual))
    return prompt, actual


# ------------------------------------------------------------- load point ---
async def run_load_point(ol, tag, prompt, opts, concurrency, num_requests, max_seconds):
    """Closed loop: `concurrency` workers pull from a shared budget until the
    budget empties or the deadline hits. Returns per-request records."""
    results = []
    budget = {"left": num_requests}
    lock = asyncio.Lock()
    deadline = time.perf_counter() + max_seconds
    limits = httpx.Limits(max_connections=concurrency + 8,
                          max_keepalive_connections=concurrency + 8)

    async with httpx.AsyncClient(timeout=httpx.Timeout(ol.timeout), limits=limits) as client:
        async def worker():
            while True:
                async with lock:
                    if budget["left"] <= 0 or time.perf_counter() > deadline:
                        return
                    budget["left"] -= 1
                results.append(await ol.generate(client, tag, prompt, opts))

        t0 = time.perf_counter()
        await asyncio.gather(*[worker() for _ in range(concurrency)])
        duration = time.perf_counter() - t0
    return results, duration


def summarise(results, duration, meta):
    ok = [r for r in results if r["ok"]]
    failed = len(results) - len(ok)
    ttft = [r["ttft_s"] * 1000 for r in ok if r["ttft_s"] is not None]
    e2e = [r["e2e_s"] * 1000 for r in ok if r["e2e_s"] is not None]
    itl_all = [x * 1000 for r in ok for x in r["itls"]]
    tpot = [
        (r["e2e_s"] - r["ttft_s"]) / (r["out_tokens"] - 1) * 1000
        for r in ok if r["ttft_s"] is not None and r["out_tokens"] > 1
    ]
    per_req_tps = [
        r["out_tokens"] / (r["e2e_s"] - r["ttft_s"])
        for r in ok if r["ttft_s"] is not None and (r["e2e_s"] - r["ttft_s"]) > 0
    ]
    tin = sum(r["in_tokens"] for r in ok)
    tout = sum(r["out_tokens"] for r in ok)
    row = dict(meta)
    row.update(
        num_requests=len(results),
        duration_s=r2(duration),
        completed_requests=len(ok),
        failed_requests=failed,
        total_input_tokens=tin,
        total_output_tokens=tout,
        ttft_p50_ms=r2(pct(ttft, 50)), ttft_p95_ms=r2(pct(ttft, 95)), ttft_p99_ms=r2(pct(ttft, 99)),
        tpot_p50_ms=r2(pct(tpot, 50), 3), tpot_p99_ms=r2(pct(tpot, 99), 3),
        itl_p50_ms=r2(pct(itl_all, 50), 3),
        e2e_p50_ms=r2(pct(e2e, 50)), e2e_p95_ms=r2(pct(e2e, 95)), e2e_p99_ms=r2(pct(e2e, 99)),
        output_tps_per_request=r2(pct(per_req_tps, 50)),
        total_output_tps=r2(tout / duration) if duration else None,
        total_throughput_tps=r2((tin + tout) / duration) if duration else None,
        requests_per_s=r2(len(ok) / duration, 3) if duration else None,
    )
    errs = [r["error"] for r in results if r.get("error")]
    if errs:
        row["notes"] = (row.get("notes", "") + f" | {len(errs)} errors, first: {errs[0]}")[:480]
    return row


# ---------------------------------------------------------- orchestration ---
def detect_env(ol):
    env = {k: os.environ.get(k, "") for k in (
        "OLLAMA_NUM_PARALLEL", "OLLAMA_MAX_LOADED_MODELS", "OLLAMA_KEEP_ALIVE",
        "OLLAMA_FLASH_ATTENTION", "OLLAMA_KV_CACHE_TYPE", "OLLAMA_SCHED_SPREAD",
        "OLLAMA_GPU_OVERHEAD")}
    try:
        v = subprocess.run(["ollama", "--version"], capture_output=True, text=True, timeout=20)
        env["ollama_version"] = v.stdout.strip() or v.stderr.strip()
    except (OSError, subprocess.SubprocessError):
        env["ollama_version"] = ""
    env["ollama_host"] = ol.host
    env["client_tool"] = f"run_ollama_bench.py (httpx {httpx.__version__}, py {platform.python_version()})"
    return env


async def bench_model(ol, model, cfg, args, sampler_interval, out_fh):
    tag = model["ollama_tag"]
    key = model["model_key"]
    rows = []

    info = ol.show(tag)
    details = info.get("details", {}) or {}
    quant = details.get("quantization_level") or "GGUF-unknown"
    num_parallel = os.environ.get("OLLAMA_NUM_PARALLEL", "") or "auto"
    stack_version = args._env.get("ollama_version", "")

    # Model-level warm-up: forces the load and gives us model_load_time_s.
    ol.unload(tag)
    load_s = None
    async with httpx.AsyncClient(timeout=httpx.Timeout(ol.timeout)) as c:
        warm = await ol.generate(c, tag, "ok", {"num_predict": 1, "temperature": 0})
        load_s = warm.get("load_s")
        if warm.get("error"):
            print(f"    ! warm-up failed: {warm['error']}")
            return rows

    for isl, osl in args._shapes:
        native = model.get("native_context_len") or 0
        num_ctx = isl + osl + 256
        if native and num_ctx > native:
            print(f"    ! skipping {isl}/{osl}: needs num_ctx {num_ctx} > native {native}")
            continue
        async with httpx.AsyncClient(timeout=httpx.Timeout(ol.timeout)) as c:
            prompt, actual_isl = await calibrate_prompt(ol, c, tag, isl, num_ctx)
        print(f"    shape {isl}/{osl}: prompt calibrated to {actual_isl} tokens (num_ctx={num_ctx})")

        opts = {
            "temperature": float(cfg.get("temperature_accuracy_runs", 0) or 0),
            "top_p": float(cfg.get("top_p", 1) or 1),
            "seed": int(cfg.get("seed", 42) or 42),
            "num_predict": osl,
            "num_ctx": num_ctx,
        }
        top_k = cfg.get("top_k")
        if top_k not in (None, "", -1):
            opts["top_k"] = int(top_k)

        for conc in args._concurrency:
            n_req = min(max(conc * args.requests_per_worker, args.min_requests), args.max_requests)
            for rep in range(1, args.repeats + 1):
                # discard warm-up requests so the batch path is hot before timing
                if args.warmup:
                    await run_load_point(ol, tag, prompt, opts, min(conc, args.warmup),
                                         min(conc, args.warmup), args.max_seconds_per_point)

                sampler = GpuSampler(sampler_interval)
                base = sampler.baseline() if sampler.available else {}
                if sampler.available:
                    sampler.start()
                results, duration = await run_load_point(
                    ol, tag, prompt, opts, conc, n_req, args.max_seconds_per_point)
                gpu = {}
                if sampler.available:
                    sampler.stop()
                    gpu = sampler.aggregate(base)

                stamp = datetime.now(timezone.utc).strftime("%H%M%S")
                meta = dict(
                    run_id=f"OLL-{key}-c{conc}-{isl}x{osl}-rep{rep}-{stamp}",
                    timestamp_utc=utc_now(),
                    model_key=key,
                    quantization=quant,
                    kv_cache_dtype=os.environ.get("OLLAMA_KV_CACHE_TYPE", "") or "f16",
                    serving_stack="Ollama",
                    stack_version=stack_version,
                    tp_size=1,                       # llama.cpp splits layers, not tensors
                    pp_size=gpu.get("gpu_count"),
                    gpu_type=gpu.get("gpu_type"),
                    gpu_count=gpu.get("gpu_count"),
                    max_model_len=num_ctx,
                    prefix_caching="on",
                    scheduler_mode=f"ollama num_parallel={num_parallel}",
                    concurrency=conc,
                    isl_tokens=actual_isl or isl,
                    osl_tokens=osl,
                    model_load_time_s=r2(load_s) if rep == 1 else None,
                    gpu_mem_peak_gb=gpu.get("gpu_mem_peak_gb"),
                    gpu_util_avg_pct=gpu.get("gpu_util_avg_pct"),
                    gpu_power_avg_w=gpu.get("gpu_power_avg_w"),
                    gpu_power_peak_w=gpu.get("gpu_power_peak_w"),
                    notes=f"ollama_tag={tag}; requested_isl={isl}; layer-split (not tensor-parallel)",
                )
                row = summarise(results, duration, meta)
                rows.append(row)
                out_fh.write(json.dumps(row) + "\n")
                out_fh.flush()
                print(f"      c={conc:<4} rep{rep}  "
                      f"out_tps={row['total_output_tps']:<8} "
                      f"ttft_p95={row['ttft_p95_ms']}ms  "
                      f"per_user_tps={row['output_tps_per_request']}  "
                      f"fail={row['failed_requests']}")
                if args.cooldown:
                    time.sleep(args.cooldown)

    if args.unload_after:
        ol.unload(tag)
    return rows


def cmd_check(ol, models):
    local = ol.local_tags()
    print(f"\n{'model_key':<22}{'ollama_tag':<26}{'status':<18}size")
    print("-" * 78)
    log = []
    for m in models:
        tag = m["ollama_tag"]
        if not tag:
            status, size = "no-tag", ""
        elif tag in local:
            status = "local"
            size = f"{local[tag].get('size', 0) / 1e9:.1f} GB"
        else:
            status, size = ol.registry_status(tag), ""
        print(f"{m['model_key']:<22}{tag or '-':<26}{status:<18}{size}")
        log.append(dict(timestamp_utc=utc_now(), model_key=m["model_key"],
                        ollama_tag=tag, status=status,
                        size_gb=round(local[tag]["size"] / 1e9, 2) if tag in local else None))
    print("-" * 78)
    runnable = [x for x in log if x["status"] in ("local", "available")]
    print(f"{len(runnable)}/{len(models)} runnable via Ollama. "
          f"'no-tag' and 'not-found' need a community GGUF tag or vLLM.\n")
    return log


def main():
    ap = argparse.ArgumentParser(description="Benchmark Ollama-served models into the workbook.")
    ap.add_argument("--workbook", default=WORKBOOK)
    ap.add_argument("--host", default=os.environ.get("OLLAMA_HOST_URL", "http://localhost:11434"))
    ap.add_argument("--models", help="comma-separated model_key list")
    ap.add_argument("--priority", help="comma-separated priorities, e.g. P0,P1")
    ap.add_argument("--concurrency", help="override Run_Config concurrency_sweep")
    ap.add_argument("--shapes", help="override Run_Config isl_osl_pairs, e.g. 2048/256,8192/1024")
    ap.add_argument("--repeats", type=int, help="override Run_Config measurement_repeats")
    ap.add_argument("--warmup", type=int, default=2, help="discarded requests before each load point")
    ap.add_argument("--requests-per-worker", type=int, default=5)
    ap.add_argument("--min-requests", type=int, default=20)
    ap.add_argument("--max-requests", type=int, default=512)
    ap.add_argument("--max-seconds-per-point", type=float, default=420.0)
    ap.add_argument("--cooldown", type=float, help="override Run_Config cooldown_between_runs_s")
    ap.add_argument("--pull", action="store_true", help="pull missing models before benchmarking")
    ap.add_argument("--check", action="store_true", help="resolve tags and exit")
    ap.add_argument("--dry-run", action="store_true", help="print the run matrix and exit")
    ap.add_argument("--no-sheet", action="store_true", help="write JSONL only")
    ap.add_argument("--unload-after", action="store_true", default=True)
    ap.add_argument("--import-only", metavar="JSONL", help="load an existing JSONL into the workbook")
    args = ap.parse_args()

    sheet_io.ensure_workbook(args.workbook)
    RESULTS_DIR.mkdir(exist_ok=True)

    if args.import_only:
        rows = [json.loads(l) for l in Path(args.import_only).read_text().splitlines() if l.strip()]
        n = sheet_io.append_perf(args.workbook, rows)
        print(f"imported {n} rows into Perf_Results")
        return 0

    cfg = sheet_io.load_run_config(args.workbook)
    only = set(parse_list(args.models)) if args.models else None
    prio = {p.upper() for p in parse_list(args.priority)} if args.priority else None
    models = sheet_io.load_models(args.workbook, priority=prio, only=only)
    if not models:
        print("No models matched. Check --models / --priority against the Models sheet.")
        return 1

    ol = Ollama(args.host)
    if not ol.alive():
        print(f"Cannot reach Ollama at {args.host}. Start it with `ollama serve`.")
        return 1

    if args.check:
        log = cmd_check(ol, models)
        if not args.no_sheet:
            sheet_io.append_pull_log(args.workbook, log)
            print(f"logged {len(log)} rows to Ollama_Pull_Log")
        return 0

    args._shapes = parse_shapes(args.shapes or cfg.get("isl_osl_pairs"))
    args._concurrency = [int(x) for x in parse_list(args.concurrency or cfg.get("concurrency_sweep"))]
    args.repeats = args.repeats or int(cfg.get("measurement_repeats", 3) or 3)
    if args.cooldown is None:
        args.cooldown = float(cfg.get("cooldown_between_runs_s", 30) or 30)
    args._env = detect_env(ol)

    np_env = os.environ.get("OLLAMA_NUM_PARALLEL")
    if np_env and args._concurrency and max(args._concurrency) > int(np_env):
        print(f"WARNING: max concurrency {max(args._concurrency)} exceeds OLLAMA_NUM_PARALLEL={np_env}. "
              f"Excess requests will queue and TTFT will measure queueing, not the model.\n")
    elif not np_env:
        print("WARNING: OLLAMA_NUM_PARALLEL is unset - Ollama picks it automatically (often 4). "
              "Set it explicitly to >= your max concurrency, and restart `ollama serve`.\n")

    runnable = [m for m in models if m["ollama_tag"]]
    skipped = [m["model_key"] for m in models if not m["ollama_tag"]]
    points = len(runnable) * len(args._shapes) * len(args._concurrency) * args.repeats
    print(f"Models: {len(runnable)} runnable, {len(skipped)} without a tag {skipped or ''}")
    print(f"Shapes: {args._shapes}")
    print(f"Concurrency: {args._concurrency}   repeats: {args.repeats}")
    print(f"=> {points} load points -> {points} rows in Perf_Results\n")
    if args.dry_run:
        for m in runnable:
            print(f"  {m['model_key']:<22} {m['ollama_tag']}")
        return 0

    pull_log = []
    if args.pull:
        local = ol.local_tags()
        for m in runnable:
            tag = m["ollama_tag"]
            if tag in local:
                pull_log.append(dict(timestamp_utc=utc_now(), model_key=m["model_key"],
                                     ollama_tag=tag, status="local"))
                continue
            print(f"  pulling {tag} ...")
            ok, secs, err = ol.pull(tag)
            info = ol.show(tag) if ok else {}
            d = info.get("details", {}) or {}
            pull_log.append(dict(
                timestamp_utc=utc_now(), model_key=m["model_key"], ollama_tag=tag,
                status="pulled" if ok else "failed",
                size_gb=round(ol.local_tags().get(tag, {}).get("size", 0) / 1e9, 2) if ok else None,
                quantization_level=d.get("quantization_level"),
                parameter_size=d.get("parameter_size"), family=d.get("family"),
                pull_seconds=round(secs, 1), error=err))
            print(f"    {'ok' if ok else 'FAILED'} in {secs:.0f}s" + (f" - {err}" if err else ""))
        runnable = [m for m in runnable
                    if any(p["model_key"] == m["model_key"] and p["status"] in ("local", "pulled")
                           for p in pull_log)]

    interval = float(cfg.get("metrics_sampling_interval_s", 1) or 1)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    out_path = RESULTS_DIR / f"perf_raw-{stamp}.jsonl"
    all_rows = []
    with out_path.open("w", encoding="utf-8") as fh:
        for m in runnable:
            print(f"\n=== {m['model_key']}  ({m['ollama_tag']}) ===")
            try:
                rows = asyncio.run(bench_model(ol, m, cfg, args, interval, fh))
                all_rows.extend(rows)
            except KeyboardInterrupt:
                print("\ninterrupted - partial results are already in the JSONL")
                break
            except Exception as e:
                print(f"  ! {m['model_key']} failed: {type(e).__name__}: {e}")

    print(f"\nraw results: {out_path}  ({len(all_rows)} rows)")
    if all_rows and not args.no_sheet:
        sheet_io.update_hardware_env(args.workbook, args._env)
        if pull_log:
            sheet_io.append_pull_log(args.workbook, pull_log)
        n = sheet_io.append_perf(args.workbook, all_rows)
        print(f"wrote {n} rows to Perf_Results in {args.workbook}")
        print("NOTE: openpyxl clears cached formula values on save. Run:\n"
              f"  python recalc_win.py {args.workbook}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
