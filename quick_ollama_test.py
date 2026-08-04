"""
Standalone quality runner for Ollama. No workbook, no openpyxl, no sheet_io - just
prompts in, scores out. Use it to iterate on the DGX before committing to a full
campaign with run_ollama_quality.py.

    python quick_ollama_test.py --list
    python quick_ollama_test.py --models qwen3:32b --suites reasoning
    python quick_ollama_test.py --models qwen3:32b,llama3.3:70b --allow-code-exec
    python quick_ollama_test.py --models qwen3:32b --repeats 3 --concurrency 4

WHAT IT SHARES WITH THE MAIN RUNNER
  graders.py, and the prompts/*.jsonl format. Scores from this script and from
  run_ollama_quality.py are identical for the same outputs, because the grading
  code is the same. What it drops is the Excel layer.

WHAT THE TIMINGS MEAN
  Ollama reports prompt_eval_duration and eval_duration per request, so the
  tok/s and prefill numbers here are real. They are NOT a substitute for
  run_ollama_bench.py: this runs one request shape at whatever concurrency you
  pass, with prompt caching on and no warm-up discipline. Treat them as a smell
  test, not as capacity numbers.

CODE EXECUTION
  The coding suite runs model-generated Python. Off unless --allow-code-exec.
  Run it in a container.
"""
from __future__ import annotations

import argparse
import asyncio
import csv
import json
import math
import statistics
import sys
import time
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

try:
    import httpx
except ImportError:
    raise SystemExit("pip install httpx")

try:
    import graders
except ImportError:
    raise SystemExit(
        "graders.py not found. Run this from the repo root; if the file is missing, "
        "restore it with:  git restore ."
    )

DEFAULT_HOST = "http://localhost:11434"
RESULTS_DIR = Path("results")
WORDS_PER_TOKEN = 0.72


# ------------------------------------------------------------- loading -----
def load_prompts(dirs, suites=None):
    """Read every *.jsonl in each directory. A record's own 'suite' field wins,
    so prompts_extra/reasoning.jsonl folds into the reasoning suite."""
    records, seen = [], {}
    for d in dirs:
        p = Path(d)
        if not p.is_dir():
            print(f"  note: {d}/ not found, skipping")
            continue
        for path in sorted(p.glob("*.jsonl")):
            for i, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
                if not line.strip():
                    continue
                try:
                    rec = json.loads(line)
                except json.JSONDecodeError as e:
                    raise SystemExit(f"{path}:{i} is not valid JSON: {e}")
                rec.setdefault("suite", path.stem)
                rid = rec.get("id") or f"{path.stem}-{i}"
                if rid in seen:
                    raise SystemExit(f"duplicate prompt id {rid!r} in {path} and {seen[rid]}")
                seen[rid] = path
                records.append(rec)
    if suites:
        records = [r for r in records if r["suite"] in suites]
    return records


# -------------------------------------------------------------- ollama -----
class Ollama:
    def __init__(self, host, timeout=600.0):
        self.host = host.rstrip("/")
        self.timeout = timeout

    async def tags(self, client):
        r = await client.get(f"{self.host}/api/tags", timeout=30.0)
        r.raise_for_status()
        return {m["name"] for m in r.json().get("models", [])}

    async def chat(self, client, model, rec, temperature, seed):
        """One non-streaming completion. Ollama's own counters give real timings."""
        max_tokens = int(rec.get("max_tokens", 512))
        approx_in = rec.get("approx_tokens") or int(len(rec["prompt"].split()) / WORDS_PER_TOKEN)
        messages = []
        if rec.get("system"):
            messages.append({"role": "system", "content": rec["system"]})
        messages.append({"role": "user", "content": rec["prompt"]})
        payload = {
            "model": model,
            "messages": messages,
            "stream": False,
            "options": {
                "temperature": temperature,
                "top_p": 1.0,
                "seed": seed,
                "num_predict": max_tokens,
                # Explicit, always. Ollama's default context is small and an
                # oversized prompt is silently truncated - you would benchmark a
                # 2k prompt believing it was 8k, and nothing in the output says so.
                "num_ctx": approx_in + max_tokens + 512,
            },
        }
        t0 = time.perf_counter()
        r = await client.post(f"{self.host}/api/chat", json=payload, timeout=self.timeout)
        wall = time.perf_counter() - t0
        if r.status_code != 200:
            raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
        d = r.json()
        ns = 1_000_000_000
        return {
            "text": (d.get("message") or {}).get("content") or "",
            "in_tokens": d.get("prompt_eval_count") or 0,
            "out_tokens": d.get("eval_count") or 0,
            "prefill_ms": round((d.get("prompt_eval_duration") or 0) / 1e6, 1),
            "decode_tps": round((d.get("eval_count") or 0) / ((d.get("eval_duration") or 1) / ns), 2),
            "wall_s": round(wall, 3),
            "truncated": d.get("done_reason") == "length",
        }


# --------------------------------------------------------------- stats -----
def ci95(scores):
    """Half-width of the 95% interval over graded items."""
    xs = [s for s in scores if s is not None]
    if len(xs) < 2:
        return None
    return round(1.96 * statistics.stdev(xs) / math.sqrt(len(xs)), 3)


def aggregate(items):
    """Per (model, suite) rollup, plus per-model overall."""
    by = defaultdict(list)
    for it in items:
        by[(it["model"], it["suite"])].append(it)
    rows = []
    for (model, suite), group in sorted(by.items()):
        scored = [g["score"] for g in group if g["score"] is not None]
        errs = [g for g in group if g["error"]]
        rows.append({
            "model": model,
            "suite": suite,
            "n_graded": len(scored),
            "n_total": len(group),
            "score_mean": round(statistics.mean(scored), 4) if scored else None,
            "ci95_halfwidth": ci95(scored),
            "refused_pct": round(100 * sum(1 for g in group if g.get("refused")) / len(group), 1),
            "truncated_pct": round(100 * sum(1 for g in group if g.get("truncated")) / len(group), 1),
            "avg_out_tokens": round(statistics.mean([g["out_tokens"] for g in group]), 1),
            "avg_decode_tps": round(statistics.mean(
                [g["decode_tps"] for g in group if g["decode_tps"]] or [0]), 2),
            "median_wall_s": round(statistics.median([g["wall_s"] for g in group if g["wall_s"]] or [0]), 2),
            "errors": len(errs),
        })
    return rows


# ----------------------------------------------------------- execution -----
async def run_model(ol, model, records, args, out_fh):
    sem = asyncio.Semaphore(args.concurrency)
    items = []
    done = 0
    total = len(records) * args.repeats

    async with httpx.AsyncClient() as client:
        async def one(rec, rep):
            nonlocal done
            async with sem:
                try:
                    res = await ol.chat(client, model, rec, args.temperature, args.seed + rep)
                    err = None
                except Exception as e:
                    res = {"text": "", "in_tokens": 0, "out_tokens": 0, "prefill_ms": None,
                           "decode_tps": None, "wall_s": None, "truncated": False}
                    err = f"{type(e).__name__}: {e}"[:300]
            verdict = ({"score": None, "detail": f"request failed: {err}"} if err else
                       graders.grade(rec, res["text"], allow_code_exec=args.allow_code_exec,
                                     code_timeout=args.code_timeout))
            item = {
                "model": model, "repeat": rep, "id": rec["id"], "suite": rec["suite"],
                "capability": rec.get("capability"), "grader": rec.get("grader"),
                "score": verdict.get("score"), "detail": verdict.get("detail"),
                "refused": verdict.get("refused"), "valid_json": verdict.get("valid_json"),
                "error": err, "output": res["text"][:2000], **{
                    k: res[k] for k in ("in_tokens", "out_tokens", "prefill_ms",
                                        "decode_tps", "wall_s", "truncated")},
            }
            out_fh.write(json.dumps(item, ensure_ascii=False) + "\n")
            out_fh.flush()
            items.append(item)
            done += 1
            print(f"    {done}/{total}", end="\r", flush=True)

        tasks = [one(rec, rep) for rep in range(args.repeats) for rec in records]
        for i in range(0, len(tasks), 20):
            await asyncio.gather(*tasks[i:i + 20])
    print(" " * 30, end="\r")
    return items


# ------------------------------------------------------------- reports -----
def print_table(rows):
    hdr = ["model", "suite", "score", "+/-95%", "n", "refuse%", "trunc%", "tok", "tok/s", "med s", "err"]
    w = [24, 22, 7, 7, 5, 8, 7, 7, 7, 7, 4]
    print("\n" + "".join(h.ljust(x) for h, x in zip(hdr, w)))
    print("-" * sum(w))
    for r in rows:
        cells = [
            r["model"][:23], r["suite"][:21],
            "-" if r["score_mean"] is None else f"{r['score_mean']:.3f}",
            "-" if r["ci95_halfwidth"] is None else f"{r['ci95_halfwidth']:.3f}",
            str(r["n_graded"]), f"{r['refused_pct']:.0f}", f"{r['truncated_pct']:.0f}",
            f"{r['avg_out_tokens']:.0f}", f"{r['avg_decode_tps']:.0f}",
            f"{r['median_wall_s']:.1f}", str(r["errors"]),
        ]
        print("".join(str(c).ljust(x) for c, x in zip(cells, w)))
    print("-" * sum(w))


def print_overall(items):
    by = defaultdict(list)
    for it in items:
        if it["score"] is not None:
            by[it["model"]].append(it["score"])
    if not by:
        return
    print("\nOverall (every graded item, all suites pooled)")
    print(f"{'model':<26}{'score':<9}{'+/-95%':<9}{'n':<6}verdict")
    print("-" * 70)
    ranked = sorted(by.items(), key=lambda kv: -statistics.mean(kv[1]))
    best = statistics.mean(ranked[0][1]) if ranked else 0
    for model, xs in ranked:
        m, h = statistics.mean(xs), ci95(xs)
        tie = h is not None and (m + h) >= best
        print(f"{model:<26}{m:<9.3f}{('-' if h is None else f'{h:.3f}'):<9}{len(xs):<6}"
              f"{'leader' if m == best else ('tied with leader' if tie else 'behind')}")
    print("-" * 70)
    print("Overlapping intervals mean the models are tied. Add prompts, not decimal places.")


def write_csv(rows, path):
    if not rows:
        return
    with path.open("w", newline="", encoding="utf-8") as fh:
        wr = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        wr.writeheader()
        wr.writerows(rows)


# ---------------------------------------------------------------- main -----
def main():
    ap = argparse.ArgumentParser(description="Standalone Ollama quality runner.")
    ap.add_argument("--host", default=DEFAULT_HOST, help="Ollama endpoint")
    ap.add_argument("--models", help="comma-separated Ollama tags, e.g. qwen3:32b,llama3.3:70b")
    ap.add_argument("--prompt-dirs", default="prompts,prompts_extra",
                    help="comma-separated directories of *.jsonl")
    ap.add_argument("--suites", help="comma-separated suite names; default all")
    ap.add_argument("--repeats", type=int, default=1)
    ap.add_argument("--concurrency", type=int, default=4)
    ap.add_argument("--temperature", type=float, default=0.0)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--limit", type=int, help="first N prompts per suite, for a fast smoke test")
    ap.add_argument("--allow-code-exec", action="store_true",
                    help="execute model-written Python for the coding suite")
    ap.add_argument("--code-timeout", type=float, default=15.0)
    ap.add_argument("--list", action="store_true", help="show suites and installed models, then exit")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    dirs = [d.strip() for d in args.prompt_dirs.split(",") if d.strip()]
    suites = {s.strip() for s in args.suites.split(",")} if args.suites else None
    records = load_prompts(dirs, suites)

    if args.limit:
        capped, count = [], defaultdict(int)
        for r in records:
            if count[r["suite"]] < args.limit:
                capped.append(r)
                count[r["suite"]] += 1
        records = capped

    ol = Ollama(args.host)
    if args.list:
        by = defaultdict(list)
        for r in records:
            by[r["suite"]].append(r)
        print(f"\n{'suite':<26}{'prompts':<10}graders")
        print("-" * 62)
        for s, rs in sorted(by.items()):
            gs = sorted({r.get("grader", "?") for r in rs})
            print(f"{s:<26}{len(rs):<10}{', '.join(gs)}")
        print("-" * 62)
        print(f"{'TOTAL':<26}{len(records)}")
        try:
            tags = asyncio.run(_tags(ol))
            print(f"\nmodels installed on {args.host}:")
            for t in sorted(tags):
                print(f"  {t}")
        except Exception as e:
            print(f"\ncould not reach Ollama at {args.host}: {e}")
        return 0

    if not records:
        raise SystemExit("No prompts matched. Check --prompt-dirs and --suites.")
    if not args.models:
        raise SystemExit("--models is required. Run --list to see what is installed.")

    models = [m.strip() for m in args.models.split(",") if m.strip()]
    total = len(models) * len(records) * args.repeats
    print(f"\n{len(models)} model(s) x {len(records)} prompts x {args.repeats} repeat(s) "
          f"= {total} requests at concurrency {args.concurrency}")

    # Dry run stays offline on purpose, so you can size a campaign from a laptop.
    if args.dry_run:
        by = defaultdict(int)
        for r in records:
            by[r["suite"]] += 1
        for s, n in sorted(by.items()):
            print(f"  {s:<26}{n}")
        return 0

    try:
        installed = asyncio.run(_tags(ol))
        missing = [m for m in models if m not in installed]
        if missing:
            print(f"\nnot installed on {args.host}: {', '.join(missing)}")
            print(f"pull them first, e.g.  ollama pull {missing[0]}")
            return 1
    except Exception as e:
        print(f"cannot reach Ollama at {args.host}: {e}")
        print("start it with  ollama serve   (remote: OLLAMA_HOST=0.0.0.0 ollama serve)")
        return 1

    has_code = any(r.get("grader") == "code_exec" for r in records)
    if has_code and not args.allow_code_exec:
        print("\nNOTE: coding prompts present but --allow-code-exec not set - they will run "
              "against the model and be recorded as ungraded, not scored zero.")

    RESULTS_DIR.mkdir(exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    raw_path = RESULTS_DIR / f"quick_raw-{stamp}.jsonl"
    all_items = []

    with raw_path.open("w", encoding="utf-8") as fh:
        for model in models:
            print(f"\n=== {model} ===")
            t0 = time.perf_counter()
            try:
                items = asyncio.run(run_model(ol, model, records, args, fh))
            except KeyboardInterrupt:
                print("\ninterrupted - raw items already on disk")
                break
            all_items.extend(items)
            rows = aggregate(items)
            print_table(rows)
            print(f"  {time.perf_counter() - t0:.0f}s")

    if not all_items:
        return 1
    rows = aggregate(all_items)
    if len(models) > 1:
        print("\n\nALL MODELS")
        print_table(rows)
    print_overall(all_items)

    csv_path = RESULTS_DIR / f"quick_scores-{stamp}.csv"
    write_csv(rows, csv_path)
    print(f"\nraw items : {raw_path}")
    print(f"scores    : {csv_path}")
    failed = [i for i in all_items if i["error"]]
    if failed:
        print(f"\n{len(failed)} request(s) failed, e.g. {failed[0]['error'][:160]}")
    trunc = [i for i in all_items if i.get("truncated")]
    if trunc:
        print(f"{len(trunc)} response(s) hit max_tokens and were cut off - those scores are "
              f"a budget problem, not a capability one.")
    return 0


async def _tags(ol):
    async with httpx.AsyncClient() as c:
        return await ol.tags(c)


if __name__ == "__main__":
    sys.exit(main())
