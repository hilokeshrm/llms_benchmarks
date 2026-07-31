"""
Generates dgx_llm_benchmark.xlsx - the data-collection workbook for benchmarking
open-weight LLMs on the DGX system.

Re-runnable: overwrites the workbook from scratch. Edit the definitions in this
file (not the .xlsx) when you want to add or rename a parameter, then re-run:

    python build_workbook.py
    python <skills>/xlsx/scripts/recalc.py dgx_llm_benchmark.xlsx
"""

from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "dgx_llm_benchmark.xlsx"

FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
IN_FILL = PatternFill("solid", fgColor="FFFF00")
EX_FILL = PatternFill("solid", fgColor="F2F2F2")
HDR_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
BOLD = Font(name=FONT, size=10, bold=True)
BODY = Font(name=FONT, size=10)
BLUE = Font(name=FONT, size=10, color="0000FF")          # hardcoded input
BLACK = Font(name=FONT, size=10, color="000000")         # formula
GREEN = Font(name=FONT, size=10, color="008000")         # cross-sheet link
GREY_IT = Font(name=FONT, size=10, italic=True, color="808080")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def header_row(ws, headers, row=1):
    """Write a styled header row and return {header: column_letter}."""
    colmap = {}
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        colmap[h] = get_column_letter(i)
    ws.row_dimensions[row].height = 34
    return colmap


def widths(ws, spec, default=14):
    for i in range(1, ws.max_column + 1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = spec.get(letter, default)


def kv_sheet(ws, title, subtitle, rows, start=4):
    """Field / Value / Unit / Notes sheet. `rows` = (field, value, unit, note).
    Returns {field: 'B<row>'} so formulas elsewhere can point at the values."""
    ws["A1"] = title
    ws["A1"].font = TITLE
    ws["A2"] = subtitle
    ws["A2"].font = GREY_IT
    hdr = header_row(ws, ["Field", "Value", "Unit", "Notes / how to obtain"], row=start - 1)
    refs = {}
    r = start
    for field, value, unit, note in rows:
        if field.startswith("##"):  # section divider
            c = ws.cell(row=r, column=1, value=field[2:].strip())
            c.font = BOLD
            c.fill = SUB_FILL
            for cc in range(2, 5):
                ws.cell(row=r, column=cc).fill = SUB_FILL
            r += 1
            continue
        ws.cell(row=r, column=1, value=field).font = BODY
        vc = ws.cell(row=r, column=2, value=value)
        vc.font = BLUE
        vc.fill = IN_FILL
        vc.border = BOX
        ws.cell(row=r, column=3, value=unit).font = BODY
        nc = ws.cell(row=r, column=4, value=note)
        nc.font = BODY
        nc.alignment = Alignment(wrap_text=True, vertical="top")
        refs[field] = f"B{r}"
        r += 1
    ws.freeze_panes = f"A{start}"
    widths(ws, {"A": 32, "B": 26, "C": 12, "D": 62})
    return refs, r, hdr


wb = Workbook()

# ---------------------------------------------------------------- README ----
ws = wb.active
ws.title = "README"
ws["A1"] = "DGX LLM Benchmark - data collection workbook"
ws["A1"].font = Font(name=FONT, size=16, bold=True, color="1F3864")
readme = [
    ("", ""),
    ("Scope", "Open-weight LLMs served locally on the DGX system. Closed models "
              "(Claude, GPT, Gemini, Grok) cannot run here - benchmark those via API in a separate workbook."),
    ("", ""),
    ("HOW TO USE", ""),
    ("1.", "Fill Hardware_Env once per system/software-stack change. Every yellow cell is an input."),
    ("2.", "Fill Run_Config once per benchmark campaign. It holds the fixed experiment controls and the SLA thresholds."),
    ("3.", "Fill Cost_Model with your real capex/power numbers. It derives $/GPU-hour, which Perf_Results uses for $/1M tokens."),
    ("4.", "Curate Models - one row per model+variant you intend to run. model_key is the join key used by every other sheet."),
    ("5.", "Perf_Results and Quality_Results are appended by the runner script. Do not type in the black formula columns."),
    ("5a.", "Ollama runs: `python run_ollama_bench.py --check` first to resolve tags, then `--priority P0 --pull`. "
            "Ollama = llama.cpp = layer-split GGUF; its throughput is well below vLLM/TensorRT-LLM on the same weights. "
            "serving_stack marks every row, so never compare an Ollama row against a vLLM row as if it were the same system."),
    ("6.", "Summary and Cost_Model roll up automatically. Refresh/recalculate after new rows land."),
    ("", ""),
    ("COLOUR LEGEND", ""),
    ("Yellow fill + blue text", "You fill this in. Hardcoded input."),
    ("Black text", "Formula - do not overwrite."),
    ("Green text", "Link to another sheet in this workbook."),
    ("Grey italic row", "EXAMPLE row showing the expected format. Delete it, or let the runner drop it (run_id starts with 'EXAMPLE')."),
    ("", ""),
    ("SHEETS", ""),
    ("Parameters", "Data dictionary - every parameter, its unit, definition and how to measure it. Read this first."),
    ("Hardware_Env", "DGX hardware + software stack versions. Reproducibility record."),
    ("Run_Config", "Fixed experiment controls: sweeps, sampling params, seeds, SLA thresholds."),
    ("Models", "Model registry: provider, country, weights, quantizations planned."),
    ("Perf_Results", "One row per load point (model x quant x TP x concurrency x ISL/OSL). Throughput, latency, GPU, power, cost."),
    ("Quality_Results", "One row per model x task x quantization. Accuracy and behaviour metrics."),
    ("Cost_Model", "Capex + power -> $/GPU-hour -> $/1M output tokens per model."),
    ("Summary", "Per-model rollup for the final comparison table."),
    ("", ""),
    ("RULES THAT KEEP THE NUMBERS COMPARABLE", ""),
    ("-", "One variable at a time. Change quantization OR TP size OR concurrency, never two at once."),
    ("-", "Discard warm-up requests; record warmup_requests in Run_Config."),
    ("-", "Report percentiles (p50/p95/p99), not means. Means hide tail latency."),
    ("-", "Same client machine, same dataset, same tokenizer for token counting across all models."),
    ("-", "Idle the GPUs between runs and record ambient conditions - thermal throttling silently skews throughput."),
    ("-", "Verify hf_repo_id in Models before pulling weights. Repo names drift between releases."),
]
r = 3
for a, b in readme:
    ca = ws.cell(row=r, column=1, value=a)
    cb = ws.cell(row=r, column=2, value=b)
    ca.font = BOLD if b == "" and a else BODY
    cb.font = BODY
    cb.alignment = Alignment(wrap_text=True, vertical="top")
    r += 1
ws.cell(row=r + 1, column=1, value=f"Generated {date.today().isoformat()} by build_workbook.py").font = GREY_IT
widths(ws, {"A": 26, "B": 110})

# ----------------------------------------------------------- Parameters ----
ws = wb.create_sheet("Parameters")
ws["A1"] = "Parameter dictionary"
ws["A1"].font = TITLE
ws["A2"] = "Every measurable parameter in this benchmark. 'Core' = collect for every run; 'Optional' = collect where relevant."
ws["A2"].font = GREY_IT
header_row(ws, ["Category", "Parameter", "Field name", "Unit", "Definition", "How to measure", "Sheet", "Priority"], row=4)

P = [
    # Category, Parameter, field, unit, definition, how, sheet, priority
    ("Identity", "Model key", "model_key", "text", "Unique join key for a model variant.", "You define it; keep it stable.", "Models", "Core"),
    ("Identity", "Provider / country", "provider / country", "text", "Lab and country of origin - the US vs non-US split.", "Model card.", "Models", "Core"),
    ("Identity", "Total / active params", "params_total_b / params_active_b", "billions", "Total weights vs params active per token (differs for MoE).", "Model card.", "Models", "Core"),
    ("Identity", "Architecture", "architecture", "text", "Dense or MoE; affects memory vs compute profile.", "Model card.", "Models", "Core"),
    ("Identity", "Native context length", "native_context_len", "tokens", "Max context the model was trained/extended to.", "Model card / config.json.", "Models", "Core"),
    ("Identity", "License", "license", "text", "Commercial-use terms. Some 'open' models restrict deployment.", "Repo LICENSE.", "Models", "Core"),

    ("Serving config", "Quantization", "quantization", "enum", "Weight precision: FP16/BF16/FP8/INT8/INT4-AWQ/INT4-GPTQ/MXFP4.", "Set at load time.", "Perf_Results", "Core"),
    ("Serving config", "KV cache dtype", "kv_cache_dtype", "enum", "Precision of the KV cache - drives max concurrency.", "Engine flag.", "Perf_Results", "Core"),
    ("Serving config", "Serving stack", "serving_stack", "enum", "vLLM / TensorRT-LLM / SGLang / NIM. Same model, very different numbers.", "Record engine + version.", "Perf_Results", "Core"),
    ("Serving config", "Tensor parallel size", "tp_size", "count", "GPUs the model is sharded across.", "Engine flag.", "Perf_Results", "Core"),
    ("Serving config", "Pipeline parallel size", "pp_size", "count", "Pipeline stages, if used.", "Engine flag.", "Perf_Results", "Optional"),
    ("Serving config", "Max model len", "max_model_len", "tokens", "Context window the engine was started with.", "Engine flag.", "Perf_Results", "Core"),
    ("Serving config", "Prefix caching", "prefix_caching", "on/off", "Automatic prefix cache enabled.", "Engine flag.", "Perf_Results", "Optional"),
    ("Serving config", "Scheduler mode", "scheduler_mode", "text", "Continuous batching / chunked prefill settings.", "Engine flag.", "Perf_Results", "Optional"),
    ("Serving config", "Model load time", "model_load_time_s", "s", "Cold start: process start to ready-to-serve.", "Engine log timestamps.", "Perf_Results", "Optional"),

    ("Load profile", "Concurrency", "concurrency", "count", "Simultaneous in-flight requests. Sweep it to find the knee.", "Client flag.", "Perf_Results", "Core"),
    ("Load profile", "Request rate", "request_rate_rps", "req/s", "Arrival rate for open-loop (Poisson) tests; blank for closed-loop.", "Client flag.", "Perf_Results", "Optional"),
    ("Load profile", "Input sequence length", "isl_tokens", "tokens", "Prompt length. Drives prefill cost.", "Dataset / synthetic generator.", "Perf_Results", "Core"),
    ("Load profile", "Output sequence length", "osl_tokens", "tokens", "Generated length. Drives decode cost.", "max_tokens + ignore_eos.", "Perf_Results", "Core"),
    ("Load profile", "Number of requests", "num_requests", "count", "Requests in the measurement window (excl. warm-up).", "Client flag.", "Perf_Results", "Core"),
    ("Load profile", "Run duration", "duration_s", "s", "Wall-clock of the measurement window.", "Client output.", "Perf_Results", "Core"),

    ("Latency", "TTFT p50 / p95 / p99", "ttft_p50_ms / p95 / p99", "ms", "Time to first token. The number users feel first.", "Client, streaming mode only.", "Perf_Results", "Core"),
    ("Latency", "TPOT p50 / p99", "tpot_p50_ms / p99", "ms", "Time per output token after the first. Perceived typing speed.", "Client.", "Perf_Results", "Core"),
    ("Latency", "Inter-token latency", "itl_p50_ms", "ms", "Gap between consecutive tokens; exposes scheduler stalls.", "Client.", "Perf_Results", "Optional"),
    ("Latency", "End-to-end p50/p95/p99", "e2e_p50_ms / p95 / p99", "ms", "Request submit to last token.", "Client.", "Perf_Results", "Core"),

    ("Throughput", "Per-request output rate", "output_tps_per_request", "tok/s", "Decode speed a single user sees.", "Client.", "Perf_Results", "Core"),
    ("Throughput", "Total output throughput", "total_output_tps", "tok/s", "System-wide generated tokens/s. The capacity number.", "Client.", "Perf_Results", "Core"),
    ("Throughput", "Total throughput", "total_throughput_tps", "tok/s", "Input + output tokens/s. Comparable across ISL/OSL mixes.", "Client.", "Perf_Results", "Core"),
    ("Throughput", "Requests per second", "requests_per_s", "req/s", "Completed requests/s at this load point.", "Client.", "Perf_Results", "Core"),
    ("Throughput", "Output tok/s per GPU", "output_tps_per_gpu", "tok/s", "Normalises throughput by GPUs used. Derived.", "Formula.", "Perf_Results", "Core"),

    ("Resource", "Peak GPU memory", "gpu_mem_peak_gb", "GB", "Highest per-GPU memory (weights + KV + activations). Sets max concurrency.", "nvidia-smi / DCGM peak sample.", "Perf_Results", "Core"),
    ("Resource", "Avg GPU utilisation", "gpu_util_avg_pct", "%", "Mean SM utilisation across the window. Low = client- or memory-bound.", "DCGM / nvidia-smi dmon.", "Perf_Results", "Core"),
    ("Resource", "Avg / peak power", "gpu_power_avg_w / peak_w", "W", "Board power summed over GPUs in use.", "DCGM DCGM_FI_DEV_POWER_USAGE.", "Perf_Results", "Core"),
    ("Resource", "Energy consumed", "energy_wh", "Wh", "avg power x duration. Derived.", "Formula.", "Perf_Results", "Core"),
    ("Resource", "Tokens per joule", "tokens_per_joule", "tok/J", "Energy efficiency. Derived.", "Formula.", "Perf_Results", "Core"),
    ("Resource", "Energy per 1k tokens", "energy_per_1k_out_tokens_wh", "Wh", "Sustainability reporting figure. Derived.", "Formula.", "Perf_Results", "Optional"),
    ("Resource", "Prefix cache hit rate", "prefix_cache_hit_pct", "%", "Share of prefill tokens served from cache.", "Engine metrics endpoint.", "Perf_Results", "Optional"),

    ("Reliability", "Completed / failed", "completed_requests / failed_requests", "count", "Successful vs errored requests in the window.", "Client.", "Perf_Results", "Core"),
    ("Reliability", "Error rate", "error_rate_pct", "%", "failed / (completed + failed). Derived.", "Formula.", "Perf_Results", "Core"),
    ("Reliability", "Meets TTFT SLA", "meets_ttft_sla", "PASS/FAIL", "ttft_p95 vs the threshold in Run_Config. Derived.", "Formula.", "Perf_Results", "Core"),
    ("Reliability", "Meets per-user TPS SLA", "meets_tps_sla", "PASS/FAIL", "output_tps_per_request vs threshold. Together these give goodput.", "Formula.", "Perf_Results", "Core"),

    ("Cost", "Cost per 1M output tokens", "cost_per_1m_out_usd", "USD", "GPU-hours consumed x $/GPU-hour, per 1M tokens. Derived.", "Formula, uses Cost_Model.", "Perf_Results", "Core"),
    ("Cost", "System capex", "system_capex_usd", "USD", "Purchase price of the DGX system.", "Invoice.", "Cost_Model", "Core"),
    ("Cost", "Amortisation period", "amortisation_years", "years", "Depreciation window.", "Finance policy.", "Cost_Model", "Core"),
    ("Cost", "Annual utilisation", "annual_utilisation", "%", "Fraction of the year the GPUs are actually serving.", "Estimate or measure.", "Cost_Model", "Core"),
    ("Cost", "System power draw", "system_power_kw", "kW", "Sustained wall power of the whole system under load.", "PDU reading.", "Cost_Model", "Core"),
    ("Cost", "PUE", "pue", "ratio", "Datacentre power usage effectiveness (1.0 = perfect).", "Facilities.", "Cost_Model", "Core"),
    ("Cost", "Electricity price", "electricity_usd_per_kwh", "USD/kWh", "Blended tariff.", "Utility bill.", "Cost_Model", "Core"),
    ("Cost", "Derived $/GPU-hour", "cost_per_gpu_hour_usd", "USD", "All-in amortised cost of one GPU for one hour. Derived.", "Formula.", "Cost_Model", "Core"),
    ("Cost", "Break-even vs API", "-", "-", "Compare cost_per_1m_out_usd against the API price of the closest closed model.", "Manual comparison.", "Summary", "Optional"),

    ("Quality", "Task accuracy", "score_mean", "varies", "Primary metric for the task suite (accuracy, pass@1, F1, ...).", "lm-eval-harness / OpenCompass / custom.", "Quality_Results", "Core"),
    ("Quality", "Score dispersion", "score_stdev / score_ci95", "same", "Spread across repeats. Without it a 1-point gap means nothing.", "Repeat runs >= 5.", "Quality_Results", "Core"),
    ("Quality", "pass@1", "pass_at_1", "%", "Coding: fraction solved on the first attempt.", "HumanEval+ / MBPP+ / SWE-bench.", "Quality_Results", "Core"),
    ("Quality", "Instruction compliance", "instruction_compliance_pct", "%", "Verifiable-constraint adherence.", "IFEval-style checker.", "Quality_Results", "Core"),
    ("Quality", "Valid JSON rate", "valid_json_rate_pct", "%", "Parses and validates against the schema.", "Schema validator.", "Quality_Results", "Core"),
    ("Quality", "Refusal rate", "refusal_rate_pct", "%", "Refusals on benign prompts - the over-refusal tax.", "Classifier + human check.", "Quality_Results", "Core"),
    ("Quality", "Hallucination rate", "hallucination_rate_pct", "%", "Unsupported claims / wrong answers on unanswerable questions.", "Grounded QA set + judge.", "Quality_Results", "Core"),
    ("Quality", "Language", "language", "text", "Per-language score. Non-US models often lead on their home languages.", "Split the suite by language.", "Quality_Results", "Core"),
    ("Quality", "Long-context retrieval", "task_name = needle_*", "%", "Retrieval accuracy at 8k/32k/128k/512k. Report the degradation curve.", "NIAH / RULER.", "Quality_Results", "Core"),
    ("Quality", "Tool-calling accuracy", "task_name = tool_*", "%", "Right tool chosen, arguments well-formed, multi-step completion.", "BFCL / custom harness.", "Quality_Results", "Core"),
    ("Quality", "Quantization delta", "delta_vs_fp16_pct", "%", "Quality lost vs the FP16 baseline of the same model. Derived.", "Formula.", "Quality_Results", "Core"),
    ("Quality", "Judge model", "judge_model", "text", "Model used for LLM-as-judge scoring. Must NOT be a contestant.", "Record it.", "Quality_Results", "Core"),
    ("Quality", "Avg output tokens", "avg_output_tokens", "tokens", "Verbosity. A model that answers in 3x the tokens costs 3x.", "Harness.", "Quality_Results", "Core"),

    ("Controls", "Temperature / top_p / top_k", "temperature / top_p / top_k", "-", "Sampling. Use 0 for accuracy runs; hold identical across models.", "Run_Config.", "Run_Config", "Core"),
    ("Controls", "Seed", "seed", "int", "Fixed seed where the engine supports it.", "Run_Config.", "Run_Config", "Core"),
    ("Controls", "Repeats", "measurement_repeats", "count", "Runs per configuration. >= 5 for anything you will quote.", "Run_Config.", "Run_Config", "Core"),
    ("Controls", "Warm-up requests", "warmup_requests", "count", "Discarded before measuring. Non-zero always.", "Run_Config.", "Run_Config", "Core"),
    ("Controls", "Tokenizer for counting", "tokenizer_for_counting", "text", "Token counts differ per tokenizer - fix one for cross-model cost maths.", "Run_Config.", "Run_Config", "Core"),
    ("Controls", "SLA thresholds", "sla_*", "ms / tok/s", "Define goodput before you run, not after you see the data.", "Run_Config.", "Run_Config", "Core"),

    ("Environment", "GPU model / count", "gpu_model / gpu_count", "text", "e.g. 8x H100 80GB SXM.", "nvidia-smi.", "Hardware_Env", "Core"),
    ("Environment", "Driver / CUDA", "nvidia_driver / cuda_version", "text", "Reproducibility. Perf shifts across driver versions.", "nvidia-smi.", "Hardware_Env", "Core"),
    ("Environment", "Engine versions", "vllm_version / trtllm_version / sglang_version", "text", "Pin them. Throughput changes release to release.", "pip show / container tag.", "Hardware_Env", "Core"),
    ("Environment", "Interconnect", "nvlink_nvswitch", "text", "NVLink/NVSwitch topology - dominates multi-GPU TP scaling.", "nvidia-smi topo -m.", "Hardware_Env", "Core"),
    ("Environment", "Power cap", "gpu_power_cap_w", "W", "Capped GPUs invalidate cross-system comparisons.", "nvidia-smi -q -d POWER.", "Hardware_Env", "Core"),
]
r = 5
for row in P:
    for i, v in enumerate(row, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.font = BODY
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = BOX
    if row[7] == "Core":
        ws.cell(row=r, column=8).font = BOLD
    r += 1
ws.freeze_panes = "A5"
ws.auto_filter.ref = f"A4:H{r - 1}"
widths(ws, {"A": 14, "B": 26, "C": 30, "D": 11, "E": 56, "F": 34, "G": 16, "H": 10})

# --------------------------------------------------------- Hardware_Env ----
ws = wb.create_sheet("Hardware_Env")
hw_rows = [
    ("## System", "", "", ""),
    ("dgx_system_model", "", "text", "e.g. DGX H100 / DGX A100 / DGX B200 / DGX Spark"),
    ("hostname", "", "text", "Node identifier."),
    ("gpu_model", "", "text", "nvidia-smi --query-gpu=name --format=csv"),
    ("gpu_count", "", "count", "GPUs physically present in the node."),
    ("gpu_memory_gb_each", "", "GB", "e.g. 80 for H100 SXM."),
    ("nvlink_nvswitch", "", "text", "nvidia-smi topo -m - record NVLink generation / NVSwitch presence."),
    ("gpu_power_cap_w", "", "W", "nvidia-smi -q -d POWER. Note if it differs from the default."),
    ("cpu_model", "", "text", "lscpu"),
    ("cpu_cores", "", "count", "Physical cores."),
    ("system_ram_gb", "", "GB", "free -g"),
    ("storage_type", "", "text", "NVMe local / NFS - affects model load time only."),
    ("network", "", "text", "InfiniBand / Ethernet, speed. Matters for multi-node only."),
    ("## Software", "", "", ""),
    ("os_version", "", "text", "cat /etc/os-release"),
    ("kernel", "", "text", "uname -r"),
    ("nvidia_driver", "", "text", "nvidia-smi --query-gpu=driver_version --format=csv"),
    ("cuda_version", "", "text", "nvcc --version or nvidia-smi header"),
    ("container_image", "", "text", "Full image tag, e.g. nvcr.io/nvidia/tritonserver:xx.yy-py3"),
    ("python_version", "", "text", "python -V"),
    ("torch_version", "", "text", "pip show torch"),
    ("vllm_version", "", "text", "pip show vllm"),
    ("trtllm_version", "", "text", "pip show tensorrt_llm"),
    ("sglang_version", "", "text", "pip show sglang"),
    ("nim_version", "", "text", "If serving via NVIDIA NIM."),
    ("dcgm_version", "", "text", "dcgmi --version - used for power/util sampling."),
    ("client_tool", "", "text", "genai-perf / vllm bench serve / custom asyncio client + version."),
    ("## Ollama runtime (auto-filled by run_ollama_bench.py)", "", "", ""),
    ("ollama_version", "", "text", "ollama --version"),
    ("ollama_host", "http://localhost:11434", "url", "Endpoint the client hits."),
    ("OLLAMA_NUM_PARALLEL", "", "count", "CRITICAL: caps real concurrency. Requests above this queue, inflating TTFT. Set it >= your max concurrency or the sweep measures queueing, not the model."),
    ("OLLAMA_MAX_LOADED_MODELS", "", "count", "Keep at 1 while benchmarking so one model owns the GPUs."),
    ("OLLAMA_KEEP_ALIVE", "", "text", "Model residency. Use 30m+ so reloads do not pollute timings."),
    ("OLLAMA_FLASH_ATTENTION", "", "0/1", "Changes throughput materially. Record it."),
    ("OLLAMA_KV_CACHE_TYPE", "", "text", "f16 / q8_0 / q4_0 - KV cache quantization."),
    ("OLLAMA_SCHED_SPREAD", "", "0/1", "1 = spread layers across all GPUs rather than packing."),
    ("OLLAMA_GPU_OVERHEAD", "", "bytes", "Reserved VRAM per GPU, if set."),
    ("## Conditions", "", "", ""),
    ("ambient_temp_c", "", "C", "Record it - thermal throttling silently caps sustained throughput."),
    ("gpus_idle_before_run", "", "Y/N", "Confirm no other workload shared the GPUs."),
    ("mig_enabled", "", "Y/N", "MIG partitioning changes everything; note the profile if on."),
    ("notes", "", "text", "Anything else that would stop someone reproducing these numbers."),
]
_, _, _ = kv_sheet(ws, "Hardware & software environment",
                   "Fill once per system or stack change. Every number in Perf_Results is only meaningful alongside this.",
                   hw_rows)

# ------------------------------------------------------------ Run_Config ---
ws = wb.create_sheet("Run_Config")
cfg_rows = [
    ("## Campaign", "", "", ""),
    ("benchmark_name", "DGX open-weight LLM benchmark v1", "text", "Free text."),
    ("owner", "", "text", "Who ran it."),
    ("date_start", "", "date", "First measurement."),
    ("date_end", "", "date", "Last measurement."),
    ("## Sweeps", "", "", ""),
    ("concurrency_sweep", "1,2,4,8,16,32,64,128", "list", "Load points. The knee between throughput and latency is the finding."),
    ("isl_osl_pairs", "128/128, 2048/256, 8192/1024, 32768/2048", "list", "Input/output token pairs. Short-prompt and long-prompt regimes behave very differently."),
    ("quantizations_under_test", "BF16, FP8, INT4-AWQ", "list", "Always include a full-precision baseline to measure the quality delta against."),
    ("tp_sizes_under_test", "1,2,4,8", "list", "Tensor-parallel sizes to sweep for scaling efficiency."),
    ("serving_stacks_under_test", "vLLM, TensorRT-LLM", "list", "Same model on two stacks is a fair and very revealing comparison."),
    ("## Sampling parameters", "", "", ""),
    ("temperature_accuracy_runs", 0, "-", "Use 0 (greedy) for accuracy. Non-zero makes runs unreproducible."),
    ("temperature_creative_runs", 0.7, "-", "Only for open-ended generation quality."),
    ("top_p", 1, "-", "Hold identical across all models."),
    ("top_k", -1, "-", "-1 / 0 = disabled, engine dependent."),
    ("max_output_tokens", 1024, "tokens", "Cap. For perf runs pair with ignore_eos so OSL is exact."),
    ("seed", 42, "int", "Where the engine supports it."),
    ("ignore_eos_for_perf", "Y", "Y/N", "Forces the exact output length - essential for clean throughput numbers."),
    ("## Measurement discipline", "", "", ""),
    ("warmup_requests", 20, "count", "Discarded. Never zero - the first requests pay compile/cache costs."),
    ("measurement_repeats", 5, "count", "Runs per configuration. Report p50 and spread across repeats."),
    ("cooldown_between_runs_s", 60, "s", "Let the GPUs return to idle temperature."),
    ("dataset_source", "ShareGPT + synthetic fixed-length", "text", "Fixed-length synthetic for clean perf; real traces for realistic mixes."),
    ("tokenizer_for_counting", "", "text", "One tokenizer for all cost maths, or per-model counts are not comparable."),
    ("metrics_sampling_interval_s", 1, "s", "DCGM/nvidia-smi polling rate for power and utilisation."),
    ("## SLA thresholds (define before running)", "", "", ""),
    ("sla_ttft_p95_ms", 500, "ms", "Interactive chat budget. Drives meets_ttft_sla."),
    ("sla_tpot_p99_ms", 50, "ms", "Per-token budget (~20 tok/s felt speed)."),
    ("sla_min_output_tps_per_user", 20, "tok/s", "Minimum acceptable per-user generation speed. Drives meets_tps_sla."),
    ("sla_max_error_rate_pct", 1, "%", "Above this the load point does not count as served."),
    ("notes", "", "text", ""),
]
cfg, _, _ = kv_sheet(ws, "Run configuration - fixed experiment controls",
                     "Set these once per campaign and do not change them mid-way. Changing a control invalidates cross-model comparison.",
                     cfg_rows)
SLA_TTFT = f"Run_Config!$B${cfg['sla_ttft_p95_ms'][1:]}"
SLA_TPS = f"Run_Config!$B${cfg['sla_min_output_tps_per_user'][1:]}"

# --------------------------------------------------------------- Models ----
ws = wb.create_sheet("Models")
ws["A1"] = "Model registry - open-weight models runnable on the DGX"
ws["A1"].font = TITLE
ws["A2"] = ("Closed models (Claude, GPT, Gemini, Grok) are API-only and deliberately absent. "
            "Verify hf_repo_id against Hugging Face before pulling - repo names drift between releases.")
ws["A2"].font = GREY_IT
mh = ["model_key", "display_name", "provider", "country", "bloc", "open_weights",
      "params_total_b", "params_active_b", "architecture", "native_context_len",
      "license", "hf_repo_id", "local_path", "ollama_tag", "modality",
      "planned_quantizations", "priority", "notes"]
header_row(ws, mh, row=4)

# Ollama library tags. Blank = no tag I could confirm exists; either find a community
# GGUF (`ollama pull hf.co/<user>/<repo>:<quant>`) or run that model on vLLM instead.
# `run_ollama_bench.py --check` resolves every tag against the registry before you commit.
OLLAMA_TAGS = {
    "deepseek-r1": "deepseek-r1:70b",
    "deepseek-v3": "deepseek-v3:671b",
    "qwen3-235b-a22b": "qwen3:235b-a22b",
    "qwen3-32b": "qwen3:32b",
    "llama-3.3-70b": "llama3.3:70b",
    "llama-4-scout": "llama4:scout",
    "gpt-oss-120b": "gpt-oss:120b",
    "gemma-3-27b": "gemma3:27b",
    "phi-4": "phi4:14b",
    "mistral-small-3": "mistral-small:24b",
    "mixtral-8x22b": "mixtral:8x22b",
    "aya-expanse-32b": "aya-expanse:32b",
    "falcon3-10b": "falcon3:10b",
}
MODELS = [
    ("deepseek-v3", "DeepSeek-V3", "DeepSeek", "China", "Non-US", "Y", 671, 37, "MoE", 128000, "DeepSeek", "deepseek-ai/DeepSeek-V3", "", "text", "BF16, FP8", "P0", "Flagship open MoE; needs multi-node or heavy quantization"),
    ("deepseek-r1", "DeepSeek-R1", "DeepSeek", "China", "Non-US", "Y", 671, 37, "MoE", 128000, "MIT", "deepseek-ai/DeepSeek-R1", "", "text", "BF16, FP8", "P0", "Reasoning; budget for long CoT output lengths"),
    ("qwen3-235b-a22b", "Qwen3-235B-A22B", "Alibaba", "China", "Non-US", "Y", 235, 22, "MoE", 128000, "Apache-2.0", "Qwen/Qwen3-235B-A22B", "", "text", "BF16, FP8, INT4-AWQ", "P0", "Strong multilingual incl. Indic"),
    ("qwen3-32b", "Qwen3-32B", "Alibaba", "China", "Non-US", "Y", 32, 32, "Dense", 128000, "Apache-2.0", "Qwen/Qwen3-32B", "", "text", "BF16, FP8, INT4-AWQ", "P0", "Fits single GPU quantized"),
    ("llama-3.3-70b", "Llama 3.3 70B Instruct", "Meta", "USA", "US", "Y", 70, 70, "Dense", 128000, "Llama 3.3 Community", "meta-llama/Llama-3.3-70B-Instruct", "", "text", "BF16, FP8, INT4-AWQ", "P0", "De-facto baseline"),
    ("llama-4-scout", "Llama 4 Scout", "Meta", "USA", "US", "Y", 109, 17, "MoE", 10000000, "Llama 4 Community", "meta-llama/Llama-4-Scout-17B-16E-Instruct", "", "text+vision", "BF16, FP8", "P1", "Very long context claim - stress it with RULER"),
    ("gpt-oss-120b", "gpt-oss-120b", "OpenAI", "USA", "US", "Y", 117, 5.1, "MoE", 128000, "Apache-2.0", "openai/gpt-oss-120b", "", "text", "MXFP4", "P0", "Ships natively in MXFP4"),
    ("gemma-3-27b", "Gemma 3 27B IT", "Google", "USA", "US", "Y", 27, 27, "Dense", 128000, "Gemma", "google/gemma-3-27b-it", "", "text+vision", "BF16, FP8, INT4-AWQ", "P1", ""),
    ("phi-4", "Phi-4", "Microsoft", "USA", "US", "Y", 14, 14, "Dense", 16000, "MIT", "microsoft/phi-4", "", "text", "BF16, FP8", "P1", "Small-model efficiency reference"),
    ("nemotron-super-49b", "Llama Nemotron Super 49B", "NVIDIA", "USA", "US", "Y", 49, 49, "Dense", 128000, "NVIDIA Open Model", "nvidia/Llama-3_3-Nemotron-Super-49B-v1", "", "text", "BF16, FP8", "P1", "NVIDIA-tuned; expect best TRT-LLM numbers"),
    ("olmo-2-32b", "OLMo 2 32B Instruct", "Allen Institute (AI2)", "USA", "US", "Y", 32, 32, "Dense", 4096, "Apache-2.0", "allenai/OLMo-2-0325-32B-Instruct", "", "text", "BF16", "P2", "Fully open data/training - reproducibility reference"),
    ("mistral-small-3", "Mistral Small 3.x Instruct", "Mistral AI", "France", "Non-US", "Y", 24, 24, "Dense", 128000, "Apache-2.0", "mistralai/Mistral-Small-24B-Instruct-2501", "", "text", "BF16, FP8, INT4-AWQ", "P0", ""),
    ("mixtral-8x22b", "Mixtral 8x22B Instruct", "Mistral AI", "France", "Non-US", "Y", 141, 39, "MoE", 65536, "Apache-2.0", "mistralai/Mixtral-8x22B-Instruct-v0.1", "", "text", "BF16, FP8", "P2", "Older MoE baseline"),
    ("glm-4.5", "GLM-4.5", "Zhipu AI / Z.ai", "China", "Non-US", "Y", 355, 32, "MoE", 128000, "MIT", "zai-org/GLM-4.5", "", "text", "BF16, FP8", "P1", "Strong agentic/tool-use scores"),
    ("kimi-k2", "Kimi K2 Instruct", "Moonshot AI", "China", "Non-US", "Y", 1000, 32, "MoE", 128000, "Modified MIT", "moonshotai/Kimi-K2-Instruct", "", "text", "FP8, INT4", "P1", "Largest open MoE - check it fits before scheduling"),
    ("minimax-m1", "MiniMax-M1", "MiniMax", "China", "Non-US", "Y", 456, 46, "MoE", 1000000, "Apache-2.0", "MiniMaxAI/MiniMax-M1-80k", "", "text", "BF16", "P2", "Long-context specialist"),
    ("internlm3-8b", "InternLM3 8B Instruct", "Shanghai AI Lab", "China", "Non-US", "Y", 8, 8, "Dense", 32768, "Apache-2.0", "internlm/internlm3-8b-instruct", "", "text", "BF16", "P2", ""),
    ("aya-expanse-32b", "Aya Expanse 32B", "Cohere", "Canada", "Non-US", "Y", 32, 32, "Dense", 128000, "CC-BY-NC", "CohereLabs/aya-expanse-32b", "", "text", "BF16", "P1", "Multilingual specialist; NC licence - research only"),
    ("falcon3-10b", "Falcon 3 10B Instruct", "TII", "UAE", "Non-US", "Y", 10, 10, "Dense", 32768, "Falcon LLM", "tiiuae/Falcon3-10B-Instruct", "", "text", "BF16, INT4-AWQ", "P2", ""),
    ("jamba-1.5-mini", "Jamba 1.5 Mini", "AI21 Labs", "Israel", "Non-US", "Y", 52, 12, "Hybrid SSM-Transformer", 256000, "Jamba Open", "ai21labs/AI21-Jamba-1.5-Mini", "", "text", "BF16", "P2", "Mamba hybrid - different memory profile, worth a row"),
    ("exaone-4-32b", "EXAONE 4.0 32B", "LG AI Research", "South Korea", "Non-US", "Y", 32, 32, "Dense", 128000, "EXAONE (NC)", "LGAI-EXAONE/EXAONE-4.0-32B", "", "text", "BF16", "P2", "Korean + English"),
    ("sarvam-m", "Sarvam-M", "Sarvam AI", "India", "Non-US", "Y", 24, 24, "Dense", 32768, "Apache-2.0", "sarvamai/sarvam-m", "", "text", "BF16", "P1", "Indic-language specialist - include Indic tasks to make this meaningful"),
    ("sea-lion-v3", "SEA-LION v3 (Gemma-based)", "AI Singapore", "Singapore", "Non-US", "Y", 9, 9, "Dense", 8192, "Gemma", "aisingapore/Gemma-SEA-LION-v3-9B-IT", "", "text", "BF16", "P2", "SE-Asian languages"),
]
r = 5
for m in MODELS:
    row_vals = list(m)
    row_vals.insert(13, OLLAMA_TAGS.get(m[0], ""))
    for i, v in enumerate(row_vals, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.font = BLUE
        c.border = BOX
        c.alignment = Alignment(vertical="top", wrap_text=(i == 18))
    r += 1
MODELS_LAST = r - 1
ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:{get_column_letter(len(mh))}{MODELS_LAST}"
widths(ws, {"A": 20, "B": 28, "C": 20, "D": 13, "E": 9, "F": 12, "G": 13, "H": 13,
            "I": 22, "J": 16, "K": 20, "L": 42, "M": 26, "N": 24, "O": 13, "P": 22,
            "Q": 9, "R": 46})
ws.cell(row=MODELS_LAST + 2, column=1,
        value="Sizes, context lengths and repo ids are as-published and must be verified before use. "
              "A blank ollama_tag means no confirmed Ollama library tag - supply a community GGUF tag "
              "(hf.co/<user>/<repo>:<quant>) or run that model on vLLM. Add rows freely - Summary reads this list.").font = GREY_IT
ws.cell(row=MODELS_LAST + 2, column=1).alignment = Alignment(wrap_text=False)

# ------------------------------------------------------ Ollama_Pull_Log ----
ws = wb.create_sheet("Ollama_Pull_Log")
ws["A1"] = "Ollama pull / resolve log"
ws["A1"].font = TITLE
ws["A2"] = ("Appended by run_ollama_bench.py. A model that failed to pull has no rows in Perf_Results - "
            "check here before concluding a model 'scored badly'.")
ws["A2"].font = GREY_IT
PULL_H = ["timestamp_utc", "model_key", "ollama_tag", "status", "size_gb",
          "quantization_level", "parameter_size", "family", "digest", "pull_seconds", "error"]
header_row(ws, PULL_H, row=4)
ws.freeze_panes = "A5"
ws.auto_filter.ref = f"A4:{get_column_letter(len(PULL_H))}404"
widths(ws, {"A": 20, "B": 20, "C": 26, "D": 12, "E": 10, "F": 18, "G": 15,
            "H": 14, "I": 22, "J": 13, "K": 60})

# ----------------------------------------------------------- Cost_Model ----
ws = wb.create_sheet("Cost_Model")
ws["A1"] = "Self-hosted cost model"
ws["A1"].font = TITLE
ws["A2"] = "Drives cost_per_1m_out_usd in Perf_Results. Replace every yellow cell with your real figures."
ws["A2"].font = GREY_IT
header_row(ws, ["Item", "Value", "Unit", "Notes"], row=4)

cost_inputs = [
    ("system_capex_usd", 350000, "USD", "Purchase price of the DGX node. Source: your invoice."),
    ("amortisation_years", 4, "years", "Depreciation window used by finance."),
    ("gpu_count_in_system", 8, "count", "GPUs in the node."),
    ("annual_utilisation", 0.70, "%", "Fraction of wall-clock the GPUs actually serve traffic."),
    ("system_power_kw", 10.2, "kW", "Sustained wall power under load. Source: PDU reading."),
    ("pue", 1.4, "ratio", "Datacentre power usage effectiveness."),
    ("electricity_usd_per_kwh", 0.09, "USD/kWh", "Blended tariff."),
    ("opex_overhead_pct", 0.15, "%", "Support contract, rack space, admin - as a fraction of capex/yr."),
]
r = 5
cref = {}
for f_, v, u, n in cost_inputs:
    ws.cell(row=r, column=1, value=f_).font = BODY
    c = ws.cell(row=r, column=2, value=v)
    c.font = BLUE
    c.fill = IN_FILL
    c.border = BOX
    c.number_format = "0.0%" if u == "%" else ("$#,##0" if u == "USD" else "0.00")
    ws.cell(row=r, column=3, value=u).font = BODY
    nc = ws.cell(row=r, column=4, value=n)
    nc.font = BODY
    nc.alignment = Alignment(wrap_text=True, vertical="top")
    cref[f_] = r
    r += 1

r += 1
ws.cell(row=r, column=1, value="Derived").font = BOLD
ws.cell(row=r, column=1).fill = SUB_FILL
for cc in range(2, 5):
    ws.cell(row=r, column=cc).fill = SUB_FILL
r += 1
derived = [
    ("annual_capex_usd", f"=B{cref['system_capex_usd']}/B{cref['amortisation_years']}", "USD/yr", "$#,##0", "Straight-line depreciation."),
    ("annual_opex_usd", f"=B{cref['system_capex_usd']}*B{cref['opex_overhead_pct']}", "USD/yr", "$#,##0", "Support and facilities."),
    ("annual_energy_kwh", f"=B{cref['system_power_kw']}*B{cref['pue']}*8760*B{cref['annual_utilisation']}", "kWh/yr", "#,##0", "Power x PUE x hours x utilisation."),
    ("annual_energy_usd", None, "USD/yr", "$#,##0", "Energy x tariff."),
    ("annual_total_usd", None, "USD/yr", "$#,##0", "Capex + opex + energy."),
    ("gpu_hours_per_year", f"=B{cref['gpu_count_in_system']}*8760*B{cref['annual_utilisation']}", "GPU-hr", "#,##0", "Billable GPU-hours actually produced."),
    ("cost_per_gpu_hour_usd", None, "USD", "$#,##0.00", "All-in. This is what Perf_Results multiplies by."),
]
first_derived = r
rows_d = {}
for name, formula, unit, fmt, note in derived:
    ws.cell(row=r, column=1, value=name).font = BODY
    rows_d[name] = r
    r += 1
r = first_derived
for name, formula, unit, fmt, note in derived:
    if name == "annual_energy_usd":
        formula = f"=B{rows_d['annual_energy_kwh']}*B{cref['electricity_usd_per_kwh']}"
    if name == "annual_total_usd":
        formula = (f"=B{rows_d['annual_capex_usd']}+B{rows_d['annual_opex_usd']}"
                   f"+B{rows_d['annual_energy_usd']}")
    if name == "cost_per_gpu_hour_usd":
        formula = f"=IFERROR(B{rows_d['annual_total_usd']}/B{rows_d['gpu_hours_per_year']},\"\")"
    c = ws.cell(row=r, column=2, value=formula)
    c.font = BLACK
    c.border = BOX
    c.number_format = fmt
    ws.cell(row=r, column=3, value=unit).font = BODY
    nc = ws.cell(row=r, column=4, value=note)
    nc.font = BODY
    nc.alignment = Alignment(wrap_text=True, vertical="top")
    r += 1
COST_PER_GPU_HR = f"Cost_Model!$B${rows_d['cost_per_gpu_hour_usd']}"
ws.cell(row=rows_d["cost_per_gpu_hour_usd"], column=2).fill = PatternFill("solid", fgColor="E2EFDA")
widths(ws, {"A": 28, "B": 18, "C": 12, "D": 60})

# --------------------------------------------------------- Perf_Results ----
ws = wb.create_sheet("Perf_Results")
PERF_IN = [
    "run_id", "timestamp_utc", "model_key", "quantization", "kv_cache_dtype",
    "serving_stack", "stack_version", "tp_size", "pp_size", "gpu_type", "gpu_count",
    "max_model_len", "prefix_caching", "scheduler_mode",
    "concurrency", "request_rate_rps", "isl_tokens", "osl_tokens", "num_requests",
    "duration_s", "completed_requests", "failed_requests",
    "total_input_tokens", "total_output_tokens",
    "ttft_p50_ms", "ttft_p95_ms", "ttft_p99_ms", "tpot_p50_ms", "tpot_p99_ms",
    "itl_p50_ms", "e2e_p50_ms", "e2e_p95_ms", "e2e_p99_ms",
    "output_tps_per_request", "total_output_tps", "total_throughput_tps", "requests_per_s",
    "gpu_mem_peak_gb", "gpu_util_avg_pct", "gpu_power_avg_w", "gpu_power_peak_w",
    "model_load_time_s", "prefix_cache_hit_pct",
]
PERF_DER = [
    "error_rate_pct", "output_tps_per_gpu", "energy_wh", "tokens_per_joule",
    "energy_per_1k_out_tokens_wh", "cost_per_1m_out_usd", "meets_ttft_sla", "meets_tps_sla",
]
PERF = PERF_IN + PERF_DER + ["notes"]
pc = header_row(ws, PERF, row=1)

PERF_FORMULAS = {
    "error_rate_pct": "=IFERROR({failed_requests}{r}/({completed_requests}{r}+{failed_requests}{r})*100,\"\")",
    "output_tps_per_gpu": "=IFERROR({total_output_tps}{r}/{gpu_count}{r},\"\")",
    "energy_wh": "=IFERROR(IF({gpu_power_avg_w}{r}=\"\",\"\",{gpu_power_avg_w}{r}*{duration_s}{r}/3600),\"\")",
    "tokens_per_joule": "=IFERROR({total_output_tokens}{r}/({gpu_power_avg_w}{r}*{duration_s}{r}),\"\")",
    "energy_per_1k_out_tokens_wh": "=IFERROR({gpu_power_avg_w}{r}*{duration_s}{r}/3600/{total_output_tokens}{r}*1000,\"\")",
    "cost_per_1m_out_usd": ("=IFERROR(" + COST_PER_GPU_HR +
                            "*{gpu_count}{r}*{duration_s}{r}/3600/{total_output_tokens}{r}*1000000,\"\")"),
    "meets_ttft_sla": "=IF(ISNUMBER({ttft_p95_ms}{r}),IF({ttft_p95_ms}{r}<=" + SLA_TTFT + ",\"PASS\",\"FAIL\"),\"\")",
    "meets_tps_sla": "=IF(ISNUMBER({output_tps_per_request}{r}),IF({output_tps_per_request}{r}>=" + SLA_TPS + ",\"PASS\",\"FAIL\"),\"\")",
}
PERF_FMT = {
    "error_rate_pct": "0.00", "output_tps_per_gpu": "#,##0.0", "energy_wh": "#,##0.0",
    "tokens_per_joule": "0.000", "energy_per_1k_out_tokens_wh": "0.000",
    "cost_per_1m_out_usd": "$#,##0.00",
}
PERF_LAST = 400
EXAMPLE_PERF = {
    "run_id": "EXAMPLE-001", "timestamp_utc": "2026-08-01T10:15:00Z", "model_key": "EXAMPLE-model",
    "quantization": "FP8", "kv_cache_dtype": "fp8", "serving_stack": "vLLM", "stack_version": "0.9.1",
    "tp_size": 2, "pp_size": 1, "gpu_type": "H100-80GB-SXM", "gpu_count": 2, "max_model_len": 32768,
    "prefix_caching": "on", "scheduler_mode": "chunked-prefill", "concurrency": 32,
    "request_rate_rps": None, "isl_tokens": 2048, "osl_tokens": 256, "num_requests": 640,
    "duration_s": 128.4, "completed_requests": 638, "failed_requests": 2,
    "total_input_tokens": 1306624, "total_output_tokens": 163328,
    "ttft_p50_ms": 312.0, "ttft_p95_ms": 480.5, "ttft_p99_ms": 640.2,
    "tpot_p50_ms": 21.4, "tpot_p99_ms": 38.9, "itl_p50_ms": 21.0,
    "e2e_p50_ms": 5790.0, "e2e_p95_ms": 7120.0, "e2e_p99_ms": 8450.0,
    "output_tps_per_request": 46.7, "total_output_tps": 1272.0,
    "total_throughput_tps": 11447.0, "requests_per_s": 4.97,
    "gpu_mem_peak_gb": 71.8, "gpu_util_avg_pct": 92.4, "gpu_power_avg_w": 1240.0,
    "gpu_power_peak_w": 1390.0, "model_load_time_s": 74.2, "prefix_cache_hit_pct": 11.3,
    "notes": "EXAMPLE ROW - shows expected formats. Delete before real runs (the runner drops run_id starting with EXAMPLE).",
}
for r_ in range(2, PERF_LAST + 1):
    for h in PERF:
        col = pc[h]
        cell = ws[f"{col}{r_}"]
        if h in PERF_FORMULAS:
            cell.value = PERF_FORMULAS[h].format(r=r_, **pc)
            cell.font = BLACK
            if h in PERF_FMT:
                cell.number_format = PERF_FMT[h]
        elif r_ == 2:
            cell.value = EXAMPLE_PERF.get(h)
            cell.font = GREY_IT
        if r_ == 2 and h not in PERF_FORMULAS:
            cell.fill = EX_FILL
ws.freeze_panes = "D2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(PERF))}{PERF_LAST}"
for i in range(1, len(PERF) + 1):
    ws.column_dimensions[get_column_letter(i)].width = 15
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 22
ws.column_dimensions[pc["notes"]].width = 60
ws[f"{pc['quantization']}1"].comment = Comment(
    "Always include a full-precision (BF16/FP16) baseline so quality delta is measurable.", "benchmark")
ws[f"{pc['ttft_p95_ms']}1"].comment = Comment(
    "TTFT is only meaningful in streaming mode. Non-streaming runs: leave blank.", "benchmark")

dv_q = DataValidation(type="list", formula1='"BF16,FP16,FP8,INT8-SQ,INT4-AWQ,INT4-GPTQ,MXFP4,NF4"', allow_blank=True)
dv_s = DataValidation(type="list", formula1='"vLLM,TensorRT-LLM,SGLang,NVIDIA-NIM,HF-TGI,llama.cpp,Ollama"', allow_blank=True)
dv_o = DataValidation(type="list", formula1='"on,off"', allow_blank=True)
for dv, col in ((dv_q, pc["quantization"]), (dv_s, pc["serving_stack"]), (dv_o, pc["prefix_caching"])):
    ws.add_data_validation(dv)
    dv.add(f"{col}3:{col}{PERF_LAST}")

# ------------------------------------------------------ Quality_Results ----
ws = wb.create_sheet("Quality_Results")
QIN = [
    "run_id", "timestamp_utc", "model_key", "quantization", "serving_stack",
    "task_suite", "task_name", "capability", "language",
    "n_samples", "repeats", "temperature", "top_p", "max_tokens", "seed",
    "metric_name", "score_mean", "score_stdev", "score_min", "score_max",
    "pass_at_1", "instruction_compliance_pct", "valid_json_rate_pct",
    "refusal_rate_pct", "hallucination_rate_pct",
    "avg_output_tokens", "avg_latency_s", "judge_model", "harness",
]
QDER = ["score_ci95_halfwidth", "fp16_baseline_score", "delta_vs_fp16_pct"]
QCOLS = QIN + QDER + ["notes"]
qc = header_row(ws, QCOLS, row=1)
QLAST = 300
QFORM = {
    "score_ci95_halfwidth": "=IFERROR(1.96*{score_stdev}{r}/SQRT({repeats}{r}),\"\")",
    "fp16_baseline_score": ("=IFERROR(AVERAGEIFS({score_mean}$2:{score_mean}$" + str(QLAST) +
                            ",{model_key}$2:{model_key}$" + str(QLAST) + ",{model_key}{r}"
                            ",{task_name}$2:{task_name}$" + str(QLAST) + ",{task_name}{r}"
                            ",{quantization}$2:{quantization}$" + str(QLAST) + ",\"BF16\"),\"\")"),
    "delta_vs_fp16_pct": ("=IFERROR(({score_mean}{r}-{fp16_baseline_score}{r})"
                          "/{fp16_baseline_score}{r}*100,\"\")"),
}
QFMT = {"score_ci95_halfwidth": "0.000", "fp16_baseline_score": "0.000", "delta_vs_fp16_pct": "0.0"}
EXAMPLE_Q = {
    "run_id": "EXAMPLE-Q-001", "timestamp_utc": "2026-08-01T14:02:00Z", "model_key": "EXAMPLE-model",
    "quantization": "BF16", "serving_stack": "vLLM", "task_suite": "MMLU-Pro", "task_name": "mmlu_pro_all",
    "capability": "knowledge", "language": "en", "n_samples": 1000, "repeats": 5, "temperature": 0,
    "top_p": 1, "max_tokens": 1024, "seed": 42, "metric_name": "accuracy", "score_mean": 0.712,
    "score_stdev": 0.008, "score_min": 0.701, "score_max": 0.722, "pass_at_1": None,
    "instruction_compliance_pct": 94.2, "valid_json_rate_pct": 99.1, "refusal_rate_pct": 0.4,
    "hallucination_rate_pct": 6.8, "avg_output_tokens": 214, "avg_latency_s": 4.9,
    "judge_model": "n/a (exact match)", "harness": "lm-eval-harness 0.4.x",
    "notes": "EXAMPLE ROW - shows expected formats. Scores stored as fractions (0.712 = 71.2%); percentages stored 0-100.",
}
for r_ in range(2, QLAST + 1):
    for h in QCOLS:
        cell = ws[f"{qc[h]}{r_}"]
        if h in QFORM:
            cell.value = QFORM[h].format(r=r_, **qc)
            cell.font = BLACK
            if h in QFMT:
                cell.number_format = QFMT[h]
        elif r_ == 2:
            cell.value = EXAMPLE_Q.get(h)
            cell.font = GREY_IT
        if r_ == 2 and h not in QFORM:
            cell.fill = EX_FILL
ws.freeze_panes = "D2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(QCOLS))}{QLAST}"
for i in range(1, len(QCOLS) + 1):
    ws.column_dimensions[get_column_letter(i)].width = 16
ws.column_dimensions["A"].width = 22
ws.column_dimensions["C"].width = 22
ws.column_dimensions[qc["notes"]].width = 60
ws[f"{qc['judge_model']}1"].comment = Comment(
    "Never use a model that is itself under test as the judge.", "benchmark")
ws[f"{qc['fp16_baseline_score']}1"].comment = Comment(
    "Averages the BF16 rows for the same model_key + task_name. Run a BF16 baseline or this stays blank.", "benchmark")

# -------------------------------------------------------------- Summary ----
ws = wb.create_sheet("Summary")
ws["A1"] = "Per-model rollup"
ws["A1"].font = TITLE
ws["A2"] = "All formulas. Model list mirrors the Models sheet; figures aggregate Perf_Results and Quality_Results."
ws["A2"].font = GREY_IT
SUM_H = ["model_key", "display_name", "provider", "country", "perf_runs",
         "best_total_output_tps", "best_output_tps_per_gpu", "ttft_p95_at_c1_ms",
         "avg_gpu_mem_peak_gb", "best_tokens_per_joule", "min_cost_per_1m_out_usd",
         "avg_error_rate_pct", "sla_pass_points", "quality_runs", "avg_quality_score",
         "avg_delta_vs_fp16_pct"]
sc = header_row(ws, SUM_H, row=4)
P_ = "Perf_Results!"
Q_ = "Quality_Results!"
PR = lambda h: f"{P_}${pc[h]}$2:${pc[h]}${PERF_LAST}"
QR = lambda h: f"{Q_}${qc[h]}$2:${qc[h]}${QLAST}"
r = 5
for i in range(len(MODELS)):
    mrow = 5 + i
    ws[f"A{r}"] = f"=IF(Models!A{mrow}=\"\",\"\",Models!A{mrow})"
    ws[f"B{r}"] = f"=IF(Models!B{mrow}=\"\",\"\",Models!B{mrow})"
    ws[f"C{r}"] = f"=IF(Models!C{mrow}=\"\",\"\",Models!C{mrow})"
    ws[f"D{r}"] = f"=IF(Models!D{mrow}=\"\",\"\",Models!D{mrow})"
    for col in "ABCD":
        ws[f"{col}{r}"].font = GREEN
    key = f"$A{r}"
    ws[f"E{r}"] = f"=COUNTIFS({PR('model_key')},{key})"
    # MAXIFS/MINIFS return 0 with no matching rows, which reads as a real measurement.
    # Gate each on the run count so an un-benchmarked model stays visibly blank.
    has = f"$E{r}=0"
    ws[f"F{r}"] = f"=IF({has},\"\",IFERROR(_xlfn.MAXIFS({PR('total_output_tps')},{PR('model_key')},{key}),\"\"))"
    ws[f"G{r}"] = f"=IF({has},\"\",IFERROR(_xlfn.MAXIFS({PR('output_tps_per_gpu')},{PR('model_key')},{key}),\"\"))"
    ws[f"H{r}"] = f"=IFERROR(AVERAGEIFS({PR('ttft_p95_ms')},{PR('model_key')},{key},{PR('concurrency')},1),\"\")"
    ws[f"I{r}"] = f"=IFERROR(AVERAGEIFS({PR('gpu_mem_peak_gb')},{PR('model_key')},{key}),\"\")"
    ws[f"J{r}"] = f"=IF({has},\"\",IFERROR(_xlfn.MAXIFS({PR('tokens_per_joule')},{PR('model_key')},{key}),\"\"))"
    ws[f"K{r}"] = f"=IF({has},\"\",IFERROR(_xlfn.MINIFS({PR('cost_per_1m_out_usd')},{PR('model_key')},{key}),\"\"))"
    ws[f"L{r}"] = f"=IFERROR(AVERAGEIFS({PR('error_rate_pct')},{PR('model_key')},{key}),\"\")"
    ws[f"M{r}"] = (f"=COUNTIFS({PR('model_key')},{key},{PR('meets_ttft_sla')},\"PASS\","
                   f"{PR('meets_tps_sla')},\"PASS\")")
    ws[f"N{r}"] = f"=COUNTIFS({QR('model_key')},{key})"
    ws[f"O{r}"] = f"=IFERROR(AVERAGEIFS({QR('score_mean')},{QR('model_key')},{key}),\"\")"
    ws[f"P{r}"] = f"=IFERROR(AVERAGEIFS({QR('delta_vs_fp16_pct')},{QR('model_key')},{key}),\"\")"
    for col in "EFGHIJKLMNOP":
        ws[f"{col}{r}"].font = BLACK
        ws[f"{col}{r}"].border = BOX
    ws[f"F{r}"].number_format = "#,##0"
    ws[f"G{r}"].number_format = "#,##0"
    ws[f"H{r}"].number_format = "#,##0"
    ws[f"I{r}"].number_format = "0.0"
    ws[f"J{r}"].number_format = "0.000"
    ws[f"K{r}"].number_format = "$#,##0.00"
    ws[f"L{r}"].number_format = "0.00"
    ws[f"O{r}"].number_format = "0.000"
    ws[f"P{r}"].number_format = "0.0"
    r += 1
ws.freeze_panes = "B5"
widths(ws, {"A": 20, "B": 28, "C": 20, "D": 13, "E": 10, "F": 15, "G": 15, "H": 15,
            "I": 15, "J": 15, "K": 16, "L": 13, "M": 13, "N": 12, "O": 14, "P": 15})
ws.cell(row=r + 1, column=1,
        value="best_* = best load point observed. ttft_p95_at_c1 = single-user latency (concurrency = 1). "
              "sla_pass_points = load points meeting BOTH SLA thresholds - i.e. usable capacity.").font = GREY_IT

for sheet in wb:
    sheet.sheet_view.showGridLines = (sheet.title in ("Perf_Results", "Quality_Results"))

wb.save(OUT)
print(f"wrote {OUT}: {len(wb.sheetnames)} sheets -> {wb.sheetnames}")
print(f"Perf_Results columns: {len(PERF)} | Quality_Results columns: {len(QCOLS)}")
