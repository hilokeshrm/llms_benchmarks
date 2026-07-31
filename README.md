# DGX LLM benchmark

Benchmarks open-weight LLMs served locally on a DGX system and writes every measurement into
`dgx_llm_benchmark.xlsx`. The first runner targets Ollama. A vLLM / TensorRT-LLM runner is the
next piece and will reuse the same workbook and the same I/O layer.

Closed models (Claude, GPT, Gemini, Grok) cannot run here. They are API-only and belong in a
separate campaign.

## Requirements

- Python 3.9+ with `openpyxl` and `httpx`
- Ollama running and reachable (`ollama serve`)
- `nvidia-smi` on PATH for power and VRAM sampling. Without it the run still works, but the
  GPU, energy and cost columns stay empty.
- LibreOffice, only for recalculating formulas after a write

```bash
pip install openpyxl httpx
```

## The files

| File | What it does |
|---|---|
| `build_workbook.py` | Generates the workbook from scratch. Edit parameters here, never in Excel. **Rebuilding destroys collected data.** |
| `run_ollama_bench.py` | Speed. Pulls models through Ollama, runs the sweep, appends rows to `Perf_Results` |
| `run_ollama_quality.py` | Accuracy, locally. Scores models against `prompts/`, appends rows to `Quality_Results` |
| `run_api_quality.py` | Accuracy, via hosted APIs. Same suites, no GPU needed |
| `providers.json` | API endpoints, rate limits, and `model_key` to provider routing |
| `graders.py` | The nine deterministic graders |
| `prompts/*.jsonl` | The test prompts, one suite per file |
| `sheet_io.py` | Reads config and models, appends results. Shared by every runner |
| `recalc_win.py` | Recalculates formulas through LibreOffice and reports any formula errors |
| `dgx_llm_benchmark.xlsx` | The data. 10 sheets |
| `results/*.jsonl` | Raw records, written before Excel is touched |

## Quick start

```bash
python build_workbook.py
```

```bash
python run_ollama_bench.py --check
```

```bash
python run_ollama_bench.py --priority P0 --concurrency 1,4,8 --shapes 2048/256 --dry-run
```

```bash
python run_ollama_bench.py --priority P0 --pull --concurrency 1,4,8,16 --shapes 128/128,2048/256,8192/1024 --repeats 3
```

```bash
python recalc_win.py dgx_llm_benchmark.xlsx
```

---

## Before you run anything

### Set `OLLAMA_NUM_PARALLEL`

This one setting decides whether your results mean anything. Ollama defaults to roughly 4
parallel slots. If you send 32 concurrent requests to a 4-slot server, 28 of them wait in a
queue, and the TTFT you record is queue time. The model looks like it collapses under load when
nothing of the sort happened.

Set it to at least your highest concurrency, then restart the server:

```bash
export OLLAMA_NUM_PARALLEL=64
```

```bash
export OLLAMA_MAX_LOADED_MODELS=1 && export OLLAMA_KEEP_ALIVE=30m && ollama serve
```

`run_ollama_bench.py` warns when the variable is unset or below your sweep, and stamps the value
into the `scheduler_mode` column of every row so you can tell later which runs were queue-limited.

### Fill three sheets by hand

Everything with a yellow fill is an input.

1. **Hardware_Env** — GPU model, driver, CUDA, interconnect, power cap. The Ollama section fills
   itself when the runner finishes.
2. **Run_Config** — sweeps, sampling parameters, seed, warm-up count, and the SLA thresholds.
   Decide the SLA numbers before you see any data, not after.
3. **Cost_Model** — capex, amortisation, power draw, PUE, electricity price. The placeholder
   numbers in there are invented. They produce a plausible `$/GPU-hour` and a meaningless
   `$/1M tokens` until you replace them.

### Check the model tags

The `ollama_tag` column in the **Models** sheet holds my best guess at each Ollama library tag.
Around half the registry has a tag I am confident in. The rest are blank, because the model is
either absent from the library or impractical as a GGUF (DeepSeek-V3 at 671B, Kimi K2 at 1T,
GLM-4.5, MiniMax-M1, Jamba, Sarvam-M, SEA-LION).

`--check` resolves each one against the registry without downloading anything:

```bash
python run_ollama_bench.py --check
```

```
model_key             ollama_tag                status            size
qwen3-32b             qwen3:32b                 local             19.4 GB
llama-3.3-70b         llama3.3:70b              available
kimi-k2                                         no-tag
```

Fix the blanks before running. Either supply a community GGUF (`hf.co/<user>/<repo>:<quant>`) or
accept that the model has to wait for the vLLM runner. Results go to the `Ollama_Pull_Log` sheet,
which is where you look when a model has no rows in `Perf_Results` and you are wondering why.

---

## How the testing works

A **load point** is one combination of model, input/output shape, and concurrency. Each repeat of
a load point produces one row in `Perf_Results`. A sweep of 7 models across 3 shapes and 4
concurrency levels at 3 repeats is 252 rows.

### 1. Prompt calibration

Ollama exposes no tokenizer endpoint, so there is no direct way to build a prompt of exactly 2048
tokens. The runner builds one from neutral filler prose, sends it with `num_predict=1`, reads the
`prompt_eval_count` that comes back, and rescales the word count. Two to four iterations land
within 3% of target.

The filler is deliberately bland, with no instructions in it. A prompt that asks the model to do
something changes how much it generates, and output length is half of what you are measuring.

The measured length goes into `isl_tokens`. The length you asked for goes into `notes`.

### 2. Context window

`num_ctx` is set explicitly to `isl + osl + 256` on every request. Ollama's default context is
small, and an 8k prompt sent to a 4k default gets **silently truncated**. You would benchmark a
2k prompt while believing you benchmarked 8k, and nothing in the output would tell you.

Shapes that exceed a model's `native_context_len` are skipped with a printed message rather than
run and quietly wrong.

### 3. Warm-up

The model is evicted with `keep_alive=0` and loaded once. That first request reports
`load_duration`, which becomes `model_load_time_s`.

Before each load point, `--warmup` requests run and are thrown away, so the batching path is warm
when timing starts. The first few requests against a cold server pay for cache allocation and
graph setup, and including them inflates TTFT by a factor that has nothing to do with the model.

### 4. The timed loop

`concurrency` async workers draw from a shared request budget until it empties or
`--max-seconds-per-point` expires, so one slow model cannot consume the whole night. Each worker
streams `POST /api/generate` and timestamps every chunk:

- first chunk carrying content gives **TTFT**
- gaps between later chunks give the **ITL** distribution
- `(e2e - ttft) / (out_tokens - 1)` gives **TPOT**
- the final `done` frame carries authoritative `prompt_eval_count` and `eval_count`

Percentiles are computed from the raw per-request lists with linear interpolation. Means are never
reported, because tail latency is the thing users actually feel.

### 5. GPU sampling

A background thread polls `nvidia-smi` once a second for utilisation, VRAM and power draw.

It takes a baseline snapshot before the run starts, then counts a GPU as active only if its memory
climbed more than 512 MB or its utilisation went above 10%. On an 8-GPU node running a model that
fits on two, you get `gpu_count=2` and power summed across those two. Without that filter the six
idle GPUs would add roughly 400 W of nothing to your energy figure and wreck `tokens_per_joule`
and `$/1M tokens`.

### 6. Results are written twice

Every row is appended to `results/perf_raw-<timestamp>.jsonl` and flushed immediately. Excel is
only opened at the very end. If you interrupt a six-hour run at hour five, nothing is lost:

```bash
python run_ollama_bench.py --import-only results/perf_raw-20260801-140000.jsonl
```

### 7. Recalculate

openpyxl writes formulas without cached values, so after any write the derived columns read as
blank to Excel until they are computed:

```bash
python recalc_win.py dgx_llm_benchmark.xlsx
```

A clean result reports `"status": "success"` with `"total_errors": 0`. That proves the formulas
evaluate. It does not prove they are right, so spot-check a row the first time.

---

## Quality testing

`Perf_Results` says how fast a model is. It says nothing about whether the model is any good.
That is `run_ollama_quality.py`, scoring models against the prompt suites in `prompts/`.

```bash
python run_ollama_quality.py --list
```

```bash
python run_ollama_quality.py --exclude-country China --suites reasoning,knowledge,extraction
```

```bash
python run_ollama_quality.py --priority P0 --allow-code-exec --needle
```

### The suites

| Suite | Prompts | Grader | Feeds |
|---|---|---|---|
| `reasoning` | 15 | `numeric`, `exact_match` | multi-step word problems, algebra, logic |
| `knowledge` | 15 | `multiple_choice` | science, computing, geography, maths |
| `coding` | 10 | `code_exec` | `pass_at_1`, via real unit tests |
| `instruction_following` | 15 | `constraint_check` | `instruction_compliance_pct` |
| `structured_output` | 12 | `json_schema` | `valid_json_rate_pct` |
| `tool_calling` | 12 | `tool_call` | correct tool and well-formed arguments |
| `extraction` | 12 | `exact_match` | grounded QA, plus `hallucination_rate_pct` |
| `refusal` | 15 | `refusal_check` | `refusal_rate_pct` on benign prompts |
| `long_context` | generated | `needle` | retrieval accuracy by length and depth |

The prompts are written from scratch rather than lifted from MMLU, GSM8K or HumanEval. Those
public sets are in the training data of nearly every model in the registry, so a high score there
partly measures who memorised the test.

### How scoring works

Every grader is deterministic. No LLM judge is involved, which means no judge bias, no API cost,
and identical scores when you re-grade the same outputs.

- `numeric` pulls the last number from the answer and compares it within a tolerance
- `multiple_choice` extracts the option letter from several common answer shapes
- `json_schema` finds the first balanced JSON object even when the model wraps it in prose or a
  code fence, then validates it against the schema and any required values
- `constraint_check` runs programmatic checks: word counts, bullet counts, forbidden words,
  casing, regex match. All constraints must hold for the prompt to score
- `tool_call` compares the chosen tool and each expected argument
- `needle` looks for a specific access code buried at a given depth in a long document
- `refusal_check` scores 1 for answering. Every prompt in that suite is benign, so a refusal is a
  false positive, not safety
- `code_exec` writes the model's function plus a fixed test suite to a temp file and runs it

### Long context

The needle suite is generated rather than stored, because a 32k-token prompt has no business
sitting in a JSONL file. Each combination of context length and needle depth becomes its own
workbook row, so `Quality_Results` shows the degradation curve directly:

```bash
python run_ollama_quality.py --models phi-4 --suites long_context --needle-lengths 4000,16000,32000,64000 --needle-depths 0.1,0.5,0.9
```

A model that scores 100% at depth 0.1 and 40% at depth 0.5 is not really a long-context model,
whatever its context window claims.

### Code execution is off by default

The `coding` suite runs Python that a model wrote. That is genuinely worth thinking about before
you enable it on a shared machine. It requires `--allow-code-exec`, and even then each candidate
runs in a subprocess with an isolated interpreter (`-I -S`), a temporary working directory and a
timeout. Run it inside a container. Without the flag those prompts still execute against the
model, and are recorded as ungraded rather than scored zero.

### How many prompts you actually need

These suites are a **screening set**, sized to eliminate obviously weak models cheaply. They are
not large enough to rank close ones.

At 15 prompts and a score near 70%, the 95% confidence interval is roughly ±22 points. At 100
prompts it is about ±9. At 600 it is about ±4. The `score_ci95_halfwidth` column computes this
for you from `score_stdev` and `n_samples × repeats`. If two models' intervals overlap, they are
tied, and no amount of decimal places in `score_mean` changes that.

Growing a suite is just appending lines to the relevant `prompts/*.jsonl`. Nothing in the code
needs to change.

## Testing through hosted APIs instead of the DGX

`run_api_quality.py` runs the same suites against hosted endpoints, so quality
testing works from any laptop with no GPU.

**Quality only.** Scores transfer because the weights are the same. Latency and throughput do
not: on a shared free tier they measure the provider's queue depth at that moment. Nothing
timing-related is written, and every row is stamped `serving_stack=API:<provider>` so it can
never be mistaken for a DGX measurement.

### Getting keys

Four keys cover all 25 models:

| Provider | Sign up at | Covers |
|---|---|---|
| Groq | console.groq.com | Llama 3.3, Llama 4 Scout, gpt-oss, Qwen3, Kimi K2 |
| OpenRouter | openrouter.ai/keys | 17 of 25 on its own, the broadest single key |
| NVIDIA | build.nvidia.com | All three Nemotron models, incl. the Super-49B that Ollama cannot run |
| Hugging Face | huggingface.co/settings/tokens | The 7 long-tail models nothing else hosts: OLMo, Falcon, Aya, EXAONE, Sarvam-M, SEA-LION, InternLM |

Copy the template and paste each key into it. `.env` is gitignored:

```bash
cp .env.example .env
```

Never commit a key, and never paste one into a chat window. The runner reads them from the
environment, never logs them, and never writes them to the workbook.

### Running it

```bash
python run_api_quality.py --list-providers
```

```bash
python run_api_quality.py --check
```

`--check` calls each provider's `/models` endpoint and reports whether every configured model id
actually exists, suggesting near matches where it does not. The ids in `providers.json` are
unverified starting points, because providers rename and retire them constantly. Run this before
every campaign.

```bash
python run_api_quality.py --dry-run --exclude-country China
```

```bash
python run_api_quality.py --exclude-country China --suites reasoning,knowledge,extraction
```

### Routing

`providers.json` maps each `model_key` to an ordered list of candidates. The runner takes the
first one whose provider has a key, so adding a key silently improves coverage and removing one
degrades gracefully. `--prefer groq,nvidia` overrides the order when you want a specific host.

Rate limiting is per provider, from the `rpm` field. Free tiers are strict, and an unthrottled
run returns a wall of 429s that looks exactly like a model scoring zero. The client retries 429s
and 5xx with exponential backoff, honours `Retry-After`, and drops `seed` or `top_p` and retries
if a provider rejects them. A 401 raises instead of being recorded as a bad score.

## Command reference

```bash
python run_ollama_bench.py --help
```

| Flag | Default | Purpose |
|---|---|---|
| `--workbook` | `dgx_llm_benchmark.xlsx` | Target workbook |
| `--host` | `http://localhost:11434` | Ollama endpoint. Point it at the DGX to drive it remotely |
| `--models` | all | Comma-separated `model_key` list |
| `--priority` | all | Filter by the priority column, e.g. `P0,P1` |
| `--country` | all | Only these countries, e.g. `USA,France` |
| `--exclude-country` | none | Skip these countries, e.g. `China` |
| `--bloc` | all | `US` or `Non-US` |
| `--concurrency` | from Run_Config | e.g. `1,4,8,16,32` |
| `--shapes` | from Run_Config | e.g. `2048/256,8192/1024` |
| `--repeats` | from Run_Config | Repeats per load point |
| `--warmup` | 2 | Discarded requests before each load point |
| `--requests-per-worker` | 5 | Requests each worker issues, so total scales with concurrency |
| `--min-requests` / `--max-requests` | 20 / 512 | Bounds on the per-load-point total |
| `--max-seconds-per-point` | 420 | Deadline that stops a pathological configuration |
| `--cooldown` | from Run_Config | Idle seconds between load points |
| `--pull` | off | Pull missing models first |
| `--check` | off | Resolve tags and exit |
| `--dry-run` | off | Print the matrix and exit |
| `--no-sheet` | off | Write JSONL only, leave Excel alone |
| `--import-only FILE` | | Load an existing JSONL into the workbook |

The same model-selection flags work on `run_ollama_quality.py`. To run everything except the
Chinese labs:

```bash
python run_ollama_quality.py --exclude-country China
```

That filters at load time and leaves the `Models` registry untouched, so dropping the filter
brings them back.

### Sizing a run

Start small. One model, one shape, two concurrency levels, one repeat, and confirm the numbers
look sane before launching the full matrix:

```bash
python run_ollama_bench.py --models qwen3-32b --concurrency 1,8 --shapes 2048/256 --repeats 1
```

Total requests for a load point is `concurrency × requests-per-worker`, clamped to
`[min-requests, max-requests]`. At concurrency 64 that is 320 requests, and at 1024 output tokens
each it is a long wait. `--max-seconds-per-point` is the safety valve.

---

## Reading the results

`Perf_Results` holds 44 measured columns and 8 derived ones. The derived columns are formulas.
Do not type over them.

| Column | Meaning |
|---|---|
| `error_rate_pct` | failed / total. Anything above `sla_max_error_rate_pct` means the load point was not really served |
| `output_tps_per_gpu` | Throughput normalised by GPUs actually used |
| `energy_wh`, `tokens_per_joule` | Energy efficiency, from sampled board power |
| `cost_per_1m_out_usd` | GPU-seconds consumed priced through `Cost_Model` |
| `meets_ttft_sla`, `meets_tps_sla` | Pass or fail against `Run_Config` thresholds |

`Summary` rolls everything up per model. The column worth putting in a report is
`sla_pass_points`, the count of load points that met both SLA thresholds. Peak throughput at
concurrency 128 is meaningless if TTFT was 9 seconds there.

`ttft_p95_at_c1_ms` is the other one to watch. That is single-user latency with no queueing, which
is what a person sitting in front of a chat window experiences.

---

## What the Ollama numbers do and do not tell you

Ollama is llama.cpp underneath. It serves GGUF weights and splits **layers** across GPUs, which is
pipeline parallelism, not tensor parallelism. Expect somewhere around 20 to 40% of what vLLM or
TensorRT-LLM achieve from the same weights on the same hardware.

The runner records this honestly rather than flattering it: `tp_size=1`, `pp_size=<gpus used>`,
`serving_stack=Ollama`. Filter on `serving_stack` before comparing anything, and never put an
Ollama row next to a vLLM row in the same chart without labelling it.

Use this path for model quality comparison, quick triage of which models are worth deeper work,
and relative ranking. Use vLLM or TensorRT-LLM for any throughput or cost number you intend to
publish or plan capacity against.

Three more things that skew results:

**`num_predict` is a ceiling, not a floor.** Ollama has no `ignore_eos`, so a model that decides
to stop early generates fewer tokens than requested. `total_output_tokens` records the real
`eval_count`, so throughput stays correct, but two models on the same nominal shape may have
produced very different amounts of text. Check that column before comparing them.

**Prompt caching is on.** Every request in a load point uses the same calibrated prompt, so after
the first one, prefill may come from cache and TTFT drops. That is realistic for chat workloads
and optimistic for anything cold-start. Vary the prompts if you need cold-prefill numbers.

**Quantization is whatever the tag ships.** `qwen3:32b` and `llama3.3:70b` are both Q4_K_M by
default, but that is a convention, not a guarantee. Pull explicit tags such as
`llama3.3:70b-instruct-q8_0` if you want a controlled precision comparison. The runner records
what it finds in the `quantization` column instead of assuming.

---

## Troubleshooting

**`Cannot reach Ollama at ...`** — start `ollama serve`, or pass `--host http://<dgx>:11434`. For
remote access the server needs `OLLAMA_HOST=0.0.0.0`.

**A model has no rows in `Perf_Results`** — check `Ollama_Pull_Log`. It either failed to pull or
had no tag. A model that never ran is not a model that scored badly.

**Derived columns look empty in Excel** — run `recalc_win.py`. openpyxl strips cached values on
save.

**GPU columns are empty** — `nvidia-smi` was not on PATH. Everything else still recorded.

**TTFT climbs steeply with concurrency and throughput plateaus early** — almost always
`OLLAMA_NUM_PARALLEL`. Check `scheduler_mode` on those rows.

**`recalc_win.py` cannot find LibreOffice** — edit the `SOFFICE` path at the top of the file. On
Linux it is usually `/usr/bin/soffice`.

**Formula errors after a manual edit** — `recalc_win.py` prints the offending cells by sheet and
coordinate.

---

## Extending it

**Adding a model.** Add a row to the `Models` sheet. `model_key` is the join key used by every
other sheet, so keep it stable. Fill `ollama_tag` and set `priority`. Nothing in the code needs to
change. The `Summary` sheet reads the first 23 model rows, so extend those formulas if you add
more than that.

**Adding a parameter.** Add the column in `build_workbook.py` and rebuild. `sheet_io.py` reads
headers from row 1 and copies derived formulas with openpyxl's `Translator`, so it picks up new
columns with no changes. Rebuilding wipes collected data, so export `Perf_Results` first, or
re-import from the JSONL afterwards.

**Adding a runner.** Import `sheet_io` and call `append_perf(path, rows)` with dicts whose keys
match the `Perf_Results` headers. Unknown keys are ignored and formula columns are never
overwritten.

## Not built yet

- `run_vllm_bench.py`, where the real DGX throughput numbers come from
- An API runner for the closed models, so self-hosted cost can be compared against API pricing
- Non-English suites. Every prompt is currently English, which structurally hides the advantage
  that Qwen, Sarvam-M, Aya and SEA-LION are built for. The `language` field is already on every
  prompt record and every workbook row, so adding a Hindi or Chinese suite needs no code change
