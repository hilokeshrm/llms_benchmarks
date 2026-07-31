"""
Scores registry models through hosted APIs instead of local Ollama, so the
quality suites can run from any laptop.

    python run_api_quality.py --list-providers            # which keys are present
    python run_api_quality.py --check                     # resolve model ids, call nothing
    python run_api_quality.py --dry-run --exclude-country China
    python run_api_quality.py --exclude-country China --suites reasoning,knowledge

KEYS
  Read from the environment, or from a .env file in the project root (gitignored).
  Never logged, never written to the workbook. One key per provider:
      GROQ_API_KEY=...
      OPENROUTER_API_KEY=...
      NVIDIA_API_KEY=...

ROUTING
  providers.json maps each model_key to an ordered list of candidates. The runner
  picks the first candidate whose provider has a key. Model id strings there are
  unverified - providers rename and retire them constantly - so --check resolves
  every id against each provider's /models endpoint before you start.

WHAT THIS IS AND IS NOT FOR
  Quality only. The scores are comparable to a local run because the weights are
  the same. Latency and throughput are NOT recorded: on a shared free tier they
  measure the provider's queue depth, not the model and not your hardware. Rows
  are stamped serving_stack=API:<provider> so they can never be mistaken for a
  DGX measurement.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import random
import sys
import time
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import httpx

import graders
import sheet_io
from run_ollama_bench import parse_list, r2, utc_now
from run_ollama_quality import aggregate, load_suites, make_needle_suite

WORKBOOK = "dgx_llm_benchmark.xlsx"
CONFIG = "providers.json"
RESULTS_DIR = Path("results")


# ------------------------------------------------------------------ keys ---
def load_dotenv(path=".env"):
    """Minimal .env reader. Real environment variables win."""
    p = Path(path)
    if not p.exists():
        return 0
    n = 0
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        k, v = k.strip(), v.strip().strip("'\"")
        if v and k not in os.environ:
            os.environ[k] = v
            n += 1
    return n


def mask(key):
    return f"{key[:4]}...{key[-4:]}" if key and len(key) > 12 else "set"


# ---------------------------------------------------------- rate limiting ---
class RateLimiter:
    """Simple async RPM gate, one per provider. Free tiers are strict, and a
    wall of 429s looks exactly like a model scoring zero."""

    def __init__(self, rpm):
        self.min_gap = 60.0 / max(rpm, 1)
        self._lock = asyncio.Lock()
        self._next = 0.0

    async def acquire(self):
        async with self._lock:
            now = time.monotonic()
            wait = self._next - now
            if wait > 0:
                await asyncio.sleep(wait)
                now = time.monotonic()
            self._next = max(now, self._next) + self.min_gap


# ------------------------------------------------------------- API client ---
class ApiError(Exception):
    pass


class Provider:
    def __init__(self, name, cfg):
        self.name = name
        self.base_url = cfg["base_url"].rstrip("/")
        self.env = cfg["env"]
        self.rpm = int(cfg.get("rpm", 30))
        self.signup = cfg.get("signup", "")
        self.limiter = RateLimiter(self.rpm)

    @property
    def key(self):
        return os.environ.get(self.env, "").strip()

    @property
    def has_key(self):
        return bool(self.key)

    def headers(self):
        h = {"Authorization": f"Bearer {self.key}", "Content-Type": "application/json"}
        if self.name == "openrouter":
            # OpenRouter asks callers to identify themselves; harmless elsewhere.
            h["HTTP-Referer"] = "https://github.com/hilokeshrm/llms_benchmarks"
            h["X-Title"] = "llms_benchmarks"
        return h

    async def list_models(self, client):
        try:
            r = await client.get(f"{self.base_url}/models", headers=self.headers(), timeout=60.0)
            if r.status_code != 200:
                return None, f"HTTP {r.status_code}"
            data = r.json()
            items = data.get("data") if isinstance(data, dict) else data
            return {str(m.get("id") or m.get("name")) for m in items if isinstance(m, dict)}, None
        except Exception as e:
            return None, f"{type(e).__name__}: {e}"

    async def chat(self, client, model_id, system, user, max_tokens, temperature, top_p,
                   seed, max_attempts=5):
        """One completion, with 429/5xx backoff and unsupported-parameter fallback."""
        messages = ([{"role": "system", "content": system}] if system else []) + \
                   [{"role": "user", "content": user}]
        payload = {"model": model_id, "messages": messages, "max_tokens": max_tokens,
                   "temperature": temperature, "top_p": top_p, "stream": False}
        if seed is not None:
            payload["seed"] = seed
        dropped = []

        for attempt in range(max_attempts):
            await self.limiter.acquire()
            t0 = time.perf_counter()
            try:
                r = await client.post(f"{self.base_url}/chat/completions",
                                      headers=self.headers(), json=payload, timeout=180.0)
            except Exception as e:
                if attempt == max_attempts - 1:
                    raise ApiError(f"{type(e).__name__}: {e}")
                await asyncio.sleep(2 ** attempt + random.random())
                continue

            if r.status_code == 200:
                data = r.json()
                choice = (data.get("choices") or [{}])[0]
                text = (choice.get("message") or {}).get("content") or ""
                usage = data.get("usage") or {}
                return {
                    "text": text,
                    "out_tokens": usage.get("completion_tokens") or 0,
                    "in_tokens": usage.get("prompt_tokens") or 0,
                    "latency_s": time.perf_counter() - t0,
                    "dropped_params": dropped,
                }

            body = r.text[:300]
            if r.status_code == 400 and payload.get("seed") is not None and "seed" in body.lower():
                payload.pop("seed"); dropped.append("seed"); continue
            if r.status_code == 400 and "top_p" in body.lower():
                payload.pop("top_p", None); dropped.append("top_p"); continue
            if r.status_code in (429, 500, 502, 503, 504) and attempt < max_attempts - 1:
                retry_after = r.headers.get("retry-after")
                delay = float(retry_after) if (retry_after or "").replace(".", "").isdigit() \
                    else 2 ** attempt + random.random()
                await asyncio.sleep(min(delay, 60.0))
                continue
            raise ApiError(f"HTTP {r.status_code}: {body}")
        raise ApiError("exhausted retries")


def load_config(path=CONFIG):
    cfg = json.loads(Path(path).read_text(encoding="utf-8"))
    providers = {n: Provider(n, c) for n, c in cfg["providers"].items()}
    return providers, cfg["models"]


async def fetch_catalogues(providers):
    """{provider_name: set_of_model_ids} for providers with a key. A provider
    whose /models call fails maps to None, meaning 'cannot verify'."""
    cats = {}
    async with httpx.AsyncClient() as client:
        for name, p in providers.items():
            if p.has_key:
                ids, _ = await p.list_models(client)
                cats[name] = ids
    return cats


def resolve(model_keys, providers, routes, prefer=None, catalogues=None):
    """model_key -> (Provider, model_id).

    Takes the first candidate whose provider has a key AND, when the provider's
    catalogue could be fetched, actually lists that model id. Without the
    catalogue check a stale id routes happily and every request 404s, which looks
    identical to a model scoring zero.
    """
    picked, unroutable = {}, []
    for key in model_keys:
        cands = routes.get(key, [])
        if prefer:
            cands = sorted(cands, key=lambda c: prefer.index(c["provider"])
                           if c["provider"] in prefer else len(prefer))
        keyed = [c for c in cands if providers.get(c["provider"])
                 and providers[c["provider"]].has_key]
        for c in keyed:
            cat = (catalogues or {}).get(c["provider"])
            if cat is not None and c["model"] not in cat:
                continue
            picked[key] = (providers[c["provider"]], c["model"])
            break
        else:
            if not cands:
                reason = "no route in providers.json"
            elif not keyed:
                reason = "no key for " + "/".join(sorted({c["provider"] for c in cands}))
            else:
                reason = ("model id not in any catalogue: " +
                          ", ".join(f"{c['provider']}:{c['model']}" for c in keyed))
            unroutable.append((key, reason))
    return picked, unroutable


# ------------------------------------------------------------- execution ---
async def score_model(prov, model_id, records, args, seed, out_fh, model_key):
    sem = asyncio.Semaphore(args.concurrency)
    items = []

    async with httpx.AsyncClient(timeout=httpx.Timeout(200.0)) as client:
        async def one(rec, rep):
            async with sem:
                try:
                    res = await prov.chat(
                        client, model_id, rec.get("system"), rec["prompt"],
                        max_tokens=int(rec.get("max_tokens", 512)),
                        temperature=args.temperature, top_p=args.top_p,
                        seed=None if args.no_seed else seed + rep)
                    err = None
                except ApiError as e:
                    res, err = {"text": "", "out_tokens": 0, "latency_s": None}, str(e)[:300]
            text = res.get("text") or ""
            verdict = graders.grade(rec, text, allow_code_exec=args.allow_code_exec,
                                    code_timeout=args.code_timeout) if not err else \
                {"score": None, "detail": f"request failed: {err}"}
            item = {
                "model_key": model_key, "quantization": "provider-managed", "repeat": rep,
                "id": rec["id"], "suite": rec["suite"], "group": rec.get("group"),
                "capability": rec.get("capability"), "language": rec.get("language", "en"),
                "grader": rec.get("grader"), "tags": rec.get("tags", []),
                "score": verdict.get("score"), "detail": verdict.get("detail"),
                "valid_json": verdict.get("valid_json"), "refused": verdict.get("refused"),
                "out_tokens": res.get("out_tokens", 0), "latency_s": r2(res.get("latency_s"), 3),
                "error": err, "provider": prov.name, "provider_model": model_id,
                "output": text[:1500],
            }
            out_fh.write(json.dumps(item) + "\n")
            items.append(item)

        tasks = [one(rec, rep) for rep in range(args.repeats) for rec in records]
        for start in range(0, len(tasks), 25):
            await asyncio.gather(*tasks[start:start + 25])
            print(f"      {min(start + 25, len(tasks))}/{len(tasks)} graded", end="\r", flush=True)
        out_fh.flush()
    print(" " * 40, end="\r")
    return items


# ------------------------------------------------------------------ main ---
async def cmd_check(providers, routes, model_keys):
    print(f"\n{'provider':<14}{'key':<8}{'models':<9}status")
    print("-" * 62)
    catalogues = {}
    async with httpx.AsyncClient() as client:
        for name, p in providers.items():
            if not p.has_key:
                print(f"{name:<14}{'-':<8}{'-':<9}no key ({p.env}) -> {p.signup}")
                continue
            ids, err = await p.list_models(client)
            catalogues[name] = ids
            print(f"{name:<14}{'yes':<8}{len(ids) if ids else '-':<9}"
                  f"{'ok' if ids else 'could not list: ' + str(err)}")

    print(f"\n{'model_key':<21}{'provider':<13}{'model id':<44}status")
    print("-" * 96)
    ok = 0
    for key in model_keys:
        cands = routes.get(key, [])
        shown = False
        for c in cands:
            p = providers.get(c["provider"])
            if not p or not p.has_key:
                continue
            cat = catalogues.get(c["provider"])
            if cat is None:
                status = "unknown (no catalogue)"
            elif c["model"] in cat:
                status, ok, shown = "AVAILABLE", ok + 1, True
            else:
                near = [m for m in cat if c["model"].split("/")[-1].lower() in m.lower()][:1]
                status = f"not found{' - close: ' + near[0] if near else ''}"
            print(f"{key:<21}{c['provider']:<13}{c['model'][:43]:<44}{status}")
            if shown:
                break
        if not cands:
            print(f"{key:<21}{'-':<13}{'-':<44}no route in providers.json")
        elif not any(providers.get(c['provider']) and providers[c['provider']].has_key
                     for c in cands):
            need = sorted({c["provider"] for c in cands})
            print(f"{key:<21}{'-':<13}{'-':<44}needs a key for {'/'.join(need)}")
    print("-" * 96)
    print(f"{ok}/{len(model_keys)} models confirmed available with the keys you have.\n")


def main():
    ap = argparse.ArgumentParser(description="Score registry models through hosted APIs.")
    ap.add_argument("--workbook", default=WORKBOOK)
    ap.add_argument("--config", default=CONFIG)
    ap.add_argument("--models"); ap.add_argument("--priority")
    ap.add_argument("--country"); ap.add_argument("--exclude-country"); ap.add_argument("--bloc")
    ap.add_argument("--suites", help="comma-separated suite names; default all")
    ap.add_argument("--prefer", help="preferred provider order, e.g. groq,nvidia,openrouter")
    ap.add_argument("--repeats", type=int, default=1)
    ap.add_argument("--concurrency", type=int, default=4)
    ap.add_argument("--temperature", type=float, default=0.0)
    ap.add_argument("--top-p", type=float, default=1.0)
    ap.add_argument("--no-seed", action="store_true", help="omit seed (some providers reject it)")
    ap.add_argument("--allow-code-exec", action="store_true")
    ap.add_argument("--code-timeout", type=float, default=15.0)
    ap.add_argument("--needle", action="store_true")
    ap.add_argument("--needle-lengths", default="4000,16000")
    ap.add_argument("--needle-depths", default="0.1,0.5,0.9")
    ap.add_argument("--no-verify", action="store_true",
                    help="skip the catalogue check and trust providers.json as written")
    ap.add_argument("--list-providers", action="store_true")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-sheet", action="store_true")
    args = ap.parse_args()

    n = load_dotenv()
    sheet_io.ensure_workbook(args.workbook)
    RESULTS_DIR.mkdir(exist_ok=True)
    providers, routes = load_config(args.config)

    if args.list_providers:
        print(f"\nloaded {n} variables from .env\n")
        print(f"{'provider':<14}{'env var':<24}{'key':<16}{'rpm':<6}signup")
        print("-" * 96)
        for name, p in providers.items():
            print(f"{name:<14}{p.env:<24}{mask(p.key) if p.has_key else '-':<16}{p.rpm:<6}{p.signup}")
        have = [n_ for n_, p in providers.items() if p.has_key]
        print("-" * 96)
        print(f"{len(have)} provider(s) with a key: {', '.join(have) or 'none'}\n")
        return 0

    only = set(parse_list(args.models)) if args.models else None
    prio = {p.upper() for p in parse_list(args.priority)} if args.priority else None
    models = sheet_io.load_models(
        args.workbook, priority=prio, only=only,
        countries=parse_list(args.country) or None,
        exclude_countries=parse_list(args.exclude_country) or None,
        blocs=parse_list(args.bloc) or None)
    keys = [m["model_key"] for m in models]
    if not keys:
        print("No models matched.")
        return 1

    if args.check:
        asyncio.run(cmd_check(providers, routes, keys))
        return 0

    prefer = parse_list(args.prefer) or None
    catalogues = None if args.no_verify else asyncio.run(fetch_catalogues(providers))
    picked, unroutable = resolve(keys, providers, routes, prefer, catalogues)
    if not picked:
        print("No model could be routed. Add a key, then re-run --list-providers.")
        for k, why in unroutable:
            print(f"  {k:<22}{why}")
        return 1

    suite_names = set(parse_list(args.suites)) if args.suites else None
    records = load_suites(suite_names)
    if args.needle or (suite_names and "long_context" in suite_names):
        records += make_needle_suite([int(x) for x in parse_list(args.needle_lengths)],
                                     [float(x) for x in parse_list(args.needle_depths)])
    if not records:
        raise SystemExit("No prompts matched.")

    print(f"\n{'model_key':<21}{'provider':<13}model id")
    print("-" * 78)
    for k, (p, mid) in picked.items():
        print(f"{k:<21}{p.name:<13}{mid}")
    if unroutable:
        print(f"\nskipped ({len(unroutable)}):")
        for k, why in unroutable:
            print(f"  {k:<22}{why}")
    total = len(picked) * len(records) * args.repeats
    slowest = min(p.rpm for p, _ in picked.values())
    print(f"\n{len(picked)} models x {len(records)} prompts x {args.repeats} repeats = "
          f"{total} requests")
    print(f"slowest provider is {slowest} rpm, so expect at least "
          f"{total / max(slowest, 1):.0f} minutes if it carries the load\n")
    if args.dry_run:
        return 0

    if any(r["grader"] == "code_exec" for r in records) and not args.allow_code_exec:
        print("NOTE: coding suite present but --allow-code-exec not set; left ungraded.\n")

    cfg = sheet_io.load_run_config(args.workbook)
    seed = int(cfg.get("seed", 42) or 42)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    raw_path = RESULTS_DIR / f"api_quality_raw-{stamp}.jsonl"
    all_rows = []

    with raw_path.open("w", encoding="utf-8") as fh:
        for key, (prov, model_id) in picked.items():
            print(f"=== {key}  ({prov.name}: {model_id}) ===")
            meta = dict(
                timestamp_utc=utc_now(), model_key=key, quantization="provider-managed",
                serving_stack=f"API:{prov.name}", temperature=args.temperature,
                top_p=args.top_p, max_tokens=None, seed=None if args.no_seed else seed,
                judge_model="none (deterministic graders)",
                harness=f"run_api_quality.py; {prov.name}:{model_id}",
            )
            try:
                items = asyncio.run(score_model(prov, model_id, records, args, seed, fh, key))
            except KeyboardInterrupt:
                print("\ninterrupted - raw items already written")
                break
            except Exception as e:
                print(f"  ! failed: {type(e).__name__}: {e}")
                continue
            rows = aggregate(items, meta, args.repeats)
            failed = [i for i in items if i.get("error")]
            for row in rows:
                row["notes"] = (row.get("notes", "") +
                                f"; provider={prov.name}; model_id={model_id}; "
                                f"latency not comparable to local runs")[:480]
            all_rows.extend(rows)
            for row in rows:
                print(f"  {row['task_name']:<28} score={row['score_mean']:<8} n={row['n_samples']}")
            if failed:
                print(f"  ! {len(failed)} requests failed, e.g. {failed[0]['error'][:120]}")

    print(f"\nraw items: {raw_path}  ({len(all_rows)} aggregated rows)")
    if all_rows and not args.no_sheet:
        print(f"wrote {sheet_io.append_quality(args.workbook, all_rows)} rows to Quality_Results")
        print(f"NOTE: run `python recalc_win.py {args.workbook}` to recompute derived columns.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
