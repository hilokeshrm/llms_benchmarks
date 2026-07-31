"""
Scores the models in dgx_llm_benchmark.xlsx against the prompt suites in prompts/
and appends the results to Quality_Results.

    python run_ollama_quality.py --list                          # show the suites
    python run_ollama_quality.py --models qwen3-32b --suites reasoning,knowledge
    python run_ollama_quality.py --priority P0 --allow-code-exec
    python run_ollama_quality.py --import-only results/quality_raw-...jsonl

SCORING
  Every grader is deterministic: exact match, numeric comparison, multiple-choice
  letter extraction, JSON schema validation, programmatic constraint checks, tool
  call comparison, needle retrieval, and unit-test execution for code. No LLM
  judge, so re-grading the same outputs always gives the same scores.

ROWS
  One row per (model, suite), or per (model, suite, group) where a suite defines
  groups - the long-context suite groups by context length and needle depth so the
  degradation curve is visible.

CODE EXECUTION
  The coding suite runs model-generated Python. It is disabled unless you pass
  --allow-code-exec, and even then it runs in a subprocess with an isolated
  interpreter, a temp working directory and a timeout. Run it in a container.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import random
import statistics
import sys
import time
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import httpx

import graders
import sheet_io
from run_ollama_bench import CORPUS, Ollama, r2, parse_list, utc_now

WORKBOOK = "dgx_llm_benchmark.xlsx"
PROMPT_DIR = Path("prompts")
RESULTS_DIR = Path("results")

# Suites whose primary metric maps onto a dedicated Quality_Results column.
METRIC_COLUMN = {
    "coding": "pass_at_1",
    "structured_output": "valid_json_rate_pct",
    "instruction_following": "instruction_compliance_pct",
    "refusal": "refusal_rate_pct",
}


# ------------------------------------------------------- prompt loading ----
def load_suites(names=None):
    if not PROMPT_DIR.is_dir():
        raise SystemExit(f"No {PROMPT_DIR}/ directory found.")
    records = []
    for path in sorted(PROMPT_DIR.glob("*.jsonl")):
        suite = path.stem
        if names and suite not in names:
            continue
        for i, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if not line.strip():
                continue
            try:
                rec = json.loads(line)
            except json.JSONDecodeError as e:
                raise SystemExit(f"{path}:{i} is not valid JSON: {e}")
            rec.setdefault("suite", suite)
            records.append(rec)
    return records


def make_needle_suite(lengths, depths, words_per_token=0.72):
    """Long-context retrieval, generated rather than stored: a 32k-token prompt
    has no business sitting in a JSONL file."""
    out = []
    for length in lengths:
        n_words = int(length * words_per_token)
        rng = random.Random(length)
        filler = [rng.choice(CORPUS) for _ in range(n_words)]
        for depth in depths:
            code = 1000 + (length // 1000) * 7 + int(depth * 100)
            needle = f"The access code for the Meridian project is {code}."
            at = min(len(filler), max(0, int(len(filler) * depth)))
            body = " ".join(filler[:at]) + " " + needle + " " + " ".join(filler[at:])
            out.append({
                "id": f"ndl-{length}-{int(depth * 100)}",
                "suite": "long_context",
                "group": f"{length // 1000}k_d{int(depth * 100)}",
                "capability": "long_context_retrieval",
                "language": "en",
                "system": "Answer using only the document provided. Reply with the number alone.",
                "prompt": f"{body}\n\nWhat is the access code for the Meridian project?",
                "expected": {"answer": str(code)},
                "grader": "needle",
                "max_tokens": 64,
                "approx_tokens": length,
                "tags": ["long_context", f"depth_{int(depth * 100)}"],
            })
    return out


# ------------------------------------------------------------- execution ---
async def ask(ol, client, tag, rec, cfg, seed):
    prompt = rec["prompt"]
    if rec.get("system"):
        prompt = f"{rec['system']}\n\n{prompt}"
    approx_in = rec.get("approx_tokens") or int(len(prompt.split()) / 0.72)
    opts = {
        "temperature": float(cfg.get("temperature_accuracy_runs", 0) or 0),
        "top_p": float(cfg.get("top_p", 1) or 1),
        "seed": seed,
        "num_predict": int(rec.get("max_tokens", 512)),
        "num_ctx": approx_in + int(rec.get("max_tokens", 512)) + 512,
    }
    t0 = time.perf_counter()
    res = await ol.generate(client, tag, prompt, opts, collect=True)
    res["latency_s"] = time.perf_counter() - t0
    return res


async def run_suite(ol, tag, records, cfg, args, seed, out_fh, model_key, quant):
    """Runs every prompt with bounded concurrency and grades as results land."""
    sem = asyncio.Semaphore(args.concurrency)
    items = []

    async with httpx.AsyncClient(timeout=httpx.Timeout(ol.timeout)) as client:
        async def one(rec, rep):
            async with sem:
                res = await ask(ol, client, tag, rec, cfg, seed + rep)
            text = res.get("text", "")
            verdict = graders.grade(rec, text,
                                    allow_code_exec=args.allow_code_exec,
                                    code_timeout=args.code_timeout)
            item = {
                "model_key": model_key, "quantization": quant, "repeat": rep,
                "id": rec["id"], "suite": rec["suite"], "group": rec.get("group"),
                "capability": rec.get("capability"), "language": rec.get("language", "en"),
                "grader": rec.get("grader"), "tags": rec.get("tags", []),
                "score": verdict.get("score"), "detail": verdict.get("detail"),
                "valid_json": verdict.get("valid_json"), "refused": verdict.get("refused"),
                "out_tokens": res.get("out_tokens", 0), "latency_s": r2(res.get("latency_s"), 3),
                "error": res.get("error"),
                "output": text[:1500],
            }
            out_fh.write(json.dumps(item) + "\n")
            items.append(item)

        tasks = [one(rec, rep) for rep in range(args.repeats) for rec in records]
        done = 0
        for chunk_start in range(0, len(tasks), 50):
            await asyncio.gather(*tasks[chunk_start:chunk_start + 50])
            done = min(chunk_start + 50, len(tasks))
            print(f"      {done}/{len(tasks)} graded", end="\r", flush=True)
        out_fh.flush()
    print(" " * 40, end="\r")
    return items


# ----------------------------------------------------------- aggregation ---
def aggregate(items, meta, repeats):
    """One workbook row per (suite, group)."""
    rows = []
    buckets = defaultdict(list)
    for it in items:
        buckets[(it["suite"], it.get("group"))].append(it)

    for (suite, group), bucket in sorted(buckets.items(), key=lambda kv: (kv[0][0], kv[0][1] or "")):
        scored = [b for b in bucket if isinstance(b["score"], (int, float))]
        if not scored:
            continue
        scores = [float(b["score"]) for b in scored]
        n_prompts = len({b["id"] for b in scored})
        per_repeat = []
        for rep in range(repeats):
            rs = [float(b["score"]) for b in scored if b["repeat"] == rep]
            if rs:
                per_repeat.append(statistics.fmean(rs))

        row = dict(meta)
        row.update(
            run_id=f"QUAL-{meta['model_key']}-{suite}{'-' + group if group else ''}"
                   f"-{datetime.now(timezone.utc).strftime('%H%M%S')}",
            task_suite=suite,
            task_name=f"{suite}/{group}" if group else suite,
            capability=sorted({b["capability"] for b in scored if b["capability"]})[0]
            if any(b["capability"] for b in scored) else suite,
            language=sorted({b["language"] for b in scored})[0],
            n_samples=n_prompts,
            repeats=repeats,
            metric_name="accuracy",
            score_mean=r2(statistics.fmean(scores), 4),
            # Spread over every graded item, so the CI in the sheet reflects
            # prompt-set size rather than near-zero greedy-decoding variance.
            score_stdev=r2(statistics.pstdev(scores), 4) if len(scores) > 1 else 0.0,
            score_min=r2(min(per_repeat), 4) if per_repeat else None,
            score_max=r2(max(per_repeat), 4) if per_repeat else None,
            avg_output_tokens=r2(statistics.fmean([b["out_tokens"] for b in scored]), 1),
            avg_latency_s=r2(statistics.fmean([b["latency_s"] or 0 for b in scored]), 3),
        )

        col = METRIC_COLUMN.get(suite)
        if col:
            row[col] = r2(statistics.fmean(scores) * 100, 2)
        if suite == "refusal":
            # The suite is entirely benign prompts, so a refusal is a false positive.
            refused = [b for b in scored if b.get("refused")]
            row["refusal_rate_pct"] = r2(len(refused) / len(scored) * 100, 2)
        if suite == "structured_output":
            valid = [b for b in scored if b.get("valid_json")]
            row["valid_json_rate_pct"] = r2(len(valid) / len(scored) * 100, 2)
        if suite == "extraction":
            # Items tagged hallucination_resistance are unanswerable from the passage.
            unans = [b for b in scored if "hallucination_resistance" in (b["tags"] or [])]
            if unans:
                wrong = [b for b in unans if float(b["score"]) < 1.0]
                row["hallucination_rate_pct"] = r2(len(wrong) / len(unans) * 100, 2)

        errs = [b for b in bucket if b.get("error")]
        skipped = [b for b in bucket if b["score"] is None]
        notes = [f"{len(scored)} graded items"]
        if errs:
            notes.append(f"{len(errs)} request errors")
        if skipped:
            notes.append(f"{len(skipped)} skipped ({skipped[0]['detail']})")
        row["notes"] = "; ".join(notes)[:480]
        rows.append(row)
    return rows


# ------------------------------------------------------------------ main ---
def main():
    ap = argparse.ArgumentParser(description="Score models on the deterministic quality suites.")
    ap.add_argument("--workbook", default=WORKBOOK)
    ap.add_argument("--host", default="http://localhost:11434")
    ap.add_argument("--models", help="comma-separated model_key list")
    ap.add_argument("--priority", help="comma-separated priorities, e.g. P0,P1")
    ap.add_argument("--country", help="only these countries, e.g. USA,France")
    ap.add_argument("--exclude-country", help="skip these countries, e.g. China")
    ap.add_argument("--bloc", help="only this bloc: US or Non-US")
    ap.add_argument("--suites", help="comma-separated suite names; default all")
    ap.add_argument("--repeats", type=int, default=1,
                    help="repeat the whole suite N times to expose engine nondeterminism")
    ap.add_argument("--concurrency", type=int, default=4, help="parallel in-flight prompts")
    ap.add_argument("--allow-code-exec", action="store_true",
                    help="execute model-generated Python for the coding suite")
    ap.add_argument("--code-timeout", type=float, default=15.0)
    ap.add_argument("--needle", action="store_true", help="add the generated long-context suite")
    ap.add_argument("--needle-lengths", default="4000,16000,32000")
    ap.add_argument("--needle-depths", default="0.1,0.5,0.9")
    ap.add_argument("--list", action="store_true", help="show suites and prompt counts, then exit")
    ap.add_argument("--no-sheet", action="store_true")
    ap.add_argument("--import-only", metavar="JSONL")
    args = ap.parse_args()

    sheet_io.ensure_workbook(args.workbook)
    RESULTS_DIR.mkdir(exist_ok=True)

    if args.import_only:
        rows = [json.loads(l) for l in Path(args.import_only).read_text().splitlines() if l.strip()]
        # A raw item file needs aggregating first; an already-aggregated file does not.
        if rows and "run_id" not in rows[0]:
            raise SystemExit("That file holds raw per-prompt items, not aggregated rows.")
        print(f"imported {sheet_io.append_quality(args.workbook, rows)} rows into Quality_Results")
        return 0

    suite_names = set(parse_list(args.suites)) if args.suites else None
    records = load_suites(suite_names)
    if args.needle or (suite_names and "long_context" in suite_names):
        records += make_needle_suite(
            [int(x) for x in parse_list(args.needle_lengths)],
            [float(x) for x in parse_list(args.needle_depths)])

    by_suite = defaultdict(list)
    for r in records:
        by_suite[r["suite"]].append(r)

    if args.list:
        print(f"\n{'suite':<24}{'prompts':<10}graders")
        print("-" * 64)
        for s, recs in sorted(by_suite.items()):
            gs = sorted({r["grader"] for r in recs})
            print(f"{s:<24}{len(recs):<10}{', '.join(gs)}")
        print("-" * 64)
        print(f"{len(records)} prompts total\n")
        print("At n=100 and a score near 70%, the 95% CI is about +/-9 points.")
        print("These suites are a screening set. Grow them before ranking close models.\n")
        return 0

    if not records:
        raise SystemExit("No prompts matched. Check --suites against prompts/*.jsonl")

    if any(r["grader"] == "code_exec" for r in records) and not args.allow_code_exec:
        print("NOTE: coding suite present but --allow-code-exec not set. "
              "Those prompts will be run and left ungraded.\n")

    cfg = sheet_io.load_run_config(args.workbook)
    seed = int(cfg.get("seed", 42) or 42)
    only = set(parse_list(args.models)) if args.models else None
    prio = {p.upper() for p in parse_list(args.priority)} if args.priority else None
    models = [m for m in sheet_io.load_models(
        args.workbook, priority=prio, only=only,
        countries=parse_list(args.country) or None,
        exclude_countries=parse_list(args.exclude_country) or None,
        blocs=parse_list(args.bloc) or None) if m["ollama_tag"]]
    if not models:
        print("No models with an ollama_tag matched.")
        return 1

    ol = Ollama(args.host)
    if not ol.alive():
        print(f"Cannot reach Ollama at {args.host}. Start it with `ollama serve`.")
        return 1

    print(f"{len(models)} models x {len(records)} prompts x {args.repeats} repeats "
          f"= {len(models) * len(records) * args.repeats} generations\n")

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    raw_path = RESULTS_DIR / f"quality_raw-{stamp}.jsonl"
    agg_path = RESULTS_DIR / f"quality_rows-{stamp}.jsonl"
    all_rows = []

    with raw_path.open("w", encoding="utf-8") as fh:
        for m in models:
            tag = m["ollama_tag"]
            print(f"=== {m['model_key']}  ({tag}) ===")
            info = ol.show(tag)
            quant = (info.get("details", {}) or {}).get("quantization_level") or "GGUF-unknown"
            meta = dict(
                timestamp_utc=utc_now(), model_key=m["model_key"], quantization=quant,
                serving_stack="Ollama", temperature=float(cfg.get("temperature_accuracy_runs", 0) or 0),
                top_p=float(cfg.get("top_p", 1) or 1), max_tokens=None, seed=seed,
                judge_model="none (deterministic graders)",
                harness=f"run_ollama_quality.py; ollama_tag={tag}",
            )
            try:
                items = asyncio.run(run_suite(ol, tag, records, cfg, args, seed, fh,
                                              m["model_key"], quant))
            except KeyboardInterrupt:
                print("\ninterrupted - raw items already written")
                break
            except Exception as e:
                print(f"  ! failed: {type(e).__name__}: {e}")
                continue
            rows = aggregate(items, meta, args.repeats)
            all_rows.extend(rows)
            for row in rows:
                extra = ""
                for c in ("pass_at_1", "valid_json_rate_pct", "instruction_compliance_pct",
                          "refusal_rate_pct", "hallucination_rate_pct"):
                    if row.get(c) is not None:
                        extra = f"  {c}={row[c]}"
                print(f"  {row['task_name']:<28} score={row['score_mean']:<8} "
                      f"n={row['n_samples']:<4}{extra}")
            ol.unload(tag)

    agg_path.write_text("\n".join(json.dumps(r) for r in all_rows), encoding="utf-8")
    print(f"\nraw items:  {raw_path}\naggregated: {agg_path}  ({len(all_rows)} rows)")

    if all_rows and not args.no_sheet:
        n = sheet_io.append_quality(args.workbook, all_rows)
        print(f"wrote {n} rows to Quality_Results in {args.workbook}")
        print(f"NOTE: run `python recalc_win.py {args.workbook}` to recompute derived columns.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
