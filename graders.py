"""Deterministic graders for the quality suites.

Every grader returns {"score": 0.0-1.0, "detail": str}. No LLM judge is involved
anywhere, so a re-run on the same outputs always produces the same scores.

A prompt record names its own grader:

    {"id": "...", "grader": "numeric", "expected": {"answer": 25}, ...}
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import jsonschema as _jsonschema
except ImportError:
    _jsonschema = None


# ------------------------------------------------------------------ utils ---
def strip_fences(text):
    """Pull the body out of a ```lang ... ``` block, if there is one."""
    m = re.search(r"```(?:[a-zA-Z0-9_+-]*)\n(.*?)```", text, re.S)
    return m.group(1).strip() if m else text.strip()


def first_json(text):
    """First balanced JSON object/array in the text. Models like to add prose."""
    body = strip_fences(text)
    for opener, closer in (("{", "}"), ("[", "]")):
        start = body.find(opener)
        if start < 0:
            continue
        depth, in_str, esc = 0, False, False
        for i, ch in enumerate(body[start:], start):
            if in_str:
                if esc:
                    esc = False
                elif ch == "\\":
                    esc = True
                elif ch == '"':
                    in_str = False
                continue
            if ch == '"':
                in_str = True
            elif ch == opener:
                depth += 1
            elif ch == closer:
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(body[start:i + 1])
                    except json.JSONDecodeError:
                        break
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        return None


def normalise(s):
    s = str(s).strip().lower()
    s = re.sub(r"[\s,]+", " ", s)
    return s.strip(" .!\"'`*")


NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")


def numbers_in(text):
    return [float(x) for x in NUM_RE.findall(str(text).replace(",", ""))]


def words(text):
    return re.findall(r"\b[\w'-]+\b", text)


REFUSAL_MARKERS = (
    "i can't", "i cannot", "i can not", "i'm unable", "i am unable", "i won't",
    "i will not", "as an ai", "i'm not able", "i am not able", "i must decline",
    "i'm sorry, but i", "i am sorry, but i", "cannot assist", "can't assist",
    "cannot help with", "can't help with", "against my guidelines",
    "not appropriate for me", "i don't feel comfortable",
)


def looks_like_refusal(text):
    head = text.strip().lower()[:400]
    return any(m in head for m in REFUSAL_MARKERS)


# --------------------------------------------------------------- graders ----
def g_exact_match(rec, out):
    accepted = rec["expected"].get("any_of") or [rec["expected"]["answer"]]
    got = normalise(out)
    tail = normalise(out.strip().splitlines()[-1]) if out.strip() else ""
    for a in accepted:
        a = normalise(a)
        if got == a or tail == a or re.search(rf"\b{re.escape(a)}\b", got):
            return {"score": 1.0, "detail": f"matched {a!r}"}
    return {"score": 0.0, "detail": f"expected one of {accepted}, got {out.strip()[-120:]!r}"}


def g_numeric(rec, out):
    want = float(rec["expected"]["answer"])
    tol = float(rec["expected"].get("tolerance", 1e-6))
    found = numbers_in(out)
    if not found:
        return {"score": 0.0, "detail": "no number in output"}
    # The answer is nearly always the last number stated.
    if abs(found[-1] - want) <= tol:
        return {"score": 1.0, "detail": f"{found[-1]}"}
    if any(abs(f - want) <= tol for f in found[-3:]):
        return {"score": 1.0, "detail": f"found {want} within last 3 numbers"}
    return {"score": 0.0, "detail": f"expected {want}, last number {found[-1]}"}


def g_multiple_choice(rec, out):
    want = str(rec["expected"]["answer"]).strip().upper()
    body = out.strip()
    pats = (r"\b(?:answer|option|choice)\s*(?:is|:)?\s*\(?([A-Z])\)?\b",
            r"^\s*\(?([A-Z])\)?[\.\):]", r"\b([A-Z])\b\s*$")
    for p in pats:
        m = re.search(p, body, re.I | re.M)
        if m:
            got = m.group(1).upper()
            return {"score": float(got == want), "detail": f"picked {got}, wanted {want}"}
    return {"score": 0.0, "detail": f"no option letter found in {body[:120]!r}"}


def _mini_validate(obj, schema, path="$"):
    """Enough of JSON Schema for these suites, used when jsonschema is absent."""
    t = schema.get("type")
    types = {"object": dict, "array": list, "string": str, "number": (int, float),
             "integer": int, "boolean": bool}
    if t and t in types:
        if t == "number" and isinstance(obj, bool):
            return [f"{path}: expected number, got boolean"]
        if not isinstance(obj, types[t]):
            return [f"{path}: expected {t}, got {type(obj).__name__}"]
    errs = []
    if t == "object":
        for k in schema.get("required", []):
            if k not in obj:
                errs.append(f"{path}: missing required '{k}'")
        for k, sub in (schema.get("properties") or {}).items():
            if k in obj:
                errs += _mini_validate(obj[k], sub, f"{path}.{k}")
    if t == "array":
        if "minItems" in schema and len(obj) < schema["minItems"]:
            errs.append(f"{path}: needs >= {schema['minItems']} items")
        if schema.get("items"):
            for i, item in enumerate(obj):
                errs += _mini_validate(item, schema["items"], f"{path}[{i}]")
    if "enum" in schema and obj not in schema["enum"]:
        errs.append(f"{path}: {obj!r} not in {schema['enum']}")
    return errs


def g_json_schema(rec, out):
    obj = first_json(out)
    if obj is None:
        return {"score": 0.0, "detail": "output is not parseable JSON", "valid_json": False}
    schema = rec["expected"]["schema"]
    if _jsonschema is not None:
        try:
            _jsonschema.validate(obj, schema)
            errs = []
        except _jsonschema.ValidationError as e:
            errs = [e.message]
    else:
        errs = _mini_validate(obj, schema)
    if errs:
        return {"score": 0.0, "detail": "; ".join(errs[:3]), "valid_json": True}
    for k, v in (rec["expected"].get("values") or {}).items():
        if normalise(obj.get(k)) != normalise(v):
            return {"score": 0.0, "detail": f"{k}: expected {v!r}, got {obj.get(k)!r}",
                    "valid_json": True}
    return {"score": 1.0, "detail": "schema ok", "valid_json": True}


CONSTRAINTS = {
    "max_words": lambda o, v: len(words(o)) <= v,
    "min_words": lambda o, v: len(words(o)) >= v,
    "exact_words": lambda o, v: len(words(o)) == v,
    "line_count": lambda o, v: len([l for l in o.strip().splitlines() if l.strip()]) == v,
    "bullet_count": lambda o, v: len(re.findall(r"^\s*[-*•]\s+", o, re.M)) == v,
    "sentence_count": lambda o, v: len([s for s in re.split(r"[.!?]+", o) if s.strip()]) == v,
    "starts_with": lambda o, v: o.strip().lower().startswith(v.lower()),
    "ends_with": lambda o, v: o.strip().rstrip(".").lower().endswith(v.lower().rstrip(".")),
    "contains_all": lambda o, v: all(x.lower() in o.lower() for x in v),
    "contains_none": lambda o, v: not any(x.lower() in o.lower() for x in v),
    "forbid_word": lambda o, v: not re.search(rf"\b{re.escape(v)}\b", o, re.I),
    "all_lowercase": lambda o, v: (o == o.lower()) == bool(v),
    "all_uppercase": lambda o, v: (o == o.upper()) == bool(v),
    "no_punctuation": lambda o, v: (not re.search(r"[.,;:!?]", o)) == bool(v),
    "matches_regex": lambda o, v: re.search(v, o.strip(), re.S | re.M) is not None,
    "is_json": lambda o, v: (first_json(o) is not None) == bool(v),
}


def g_constraint_check(rec, out):
    """All-or-nothing per prompt, then reported as a compliance percentage."""
    checks = rec["expected"]["constraints"]
    failed = []
    for name, val in checks.items():
        fn = CONSTRAINTS.get(name)
        if fn is None:
            failed.append(f"unknown constraint {name}")
            continue
        try:
            if not fn(out, val):
                failed.append(f"{name}={val!r}")
        except Exception as e:
            failed.append(f"{name} raised {type(e).__name__}")
    if failed:
        return {"score": 0.0, "detail": "violated: " + ", ".join(failed),
                "constraints_passed": len(checks) - len(failed), "constraints_total": len(checks)}
    return {"score": 1.0, "detail": f"all {len(checks)} constraints held",
            "constraints_passed": len(checks), "constraints_total": len(checks)}


def g_tool_call(rec, out):
    obj = first_json(out)
    if obj is None:
        return {"score": 0.0, "detail": "no JSON tool call emitted"}
    if isinstance(obj, list):
        obj = obj[0] if obj else {}
    name = obj.get("name") or obj.get("tool") or obj.get("function")
    if isinstance(name, dict):
        name = name.get("name")
    args = obj.get("arguments") or obj.get("args") or obj.get("parameters") or {}
    if isinstance(args, str):
        try:
            args = json.loads(args)
        except json.JSONDecodeError:
            args = {}
    exp = rec["expected"]
    if normalise(name) != normalise(exp["tool"]):
        return {"score": 0.0, "detail": f"called {name!r}, expected {exp['tool']!r}"}
    for k, v in (exp.get("arguments") or {}).items():
        if k not in args:
            return {"score": 0.0, "detail": f"missing argument {k!r}"}
        if v is not None and normalise(args[k]) != normalise(v):
            return {"score": 0.0, "detail": f"{k}: expected {v!r}, got {args[k]!r}"}
    return {"score": 1.0, "detail": f"{name} with correct arguments"}


def g_needle(rec, out):
    want = str(rec["expected"]["answer"])
    if re.search(rf"\b{re.escape(want)}\b", out, re.I):
        return {"score": 1.0, "detail": f"retrieved {want!r}"}
    return {"score": 0.0, "detail": f"{want!r} not found in {len(out)}-char answer"}


def g_refusal_check(rec, out):
    """Benign prompts. Answering scores 1, refusing scores 0."""
    if looks_like_refusal(out):
        return {"score": 0.0, "detail": "refused a benign request", "refused": True}
    if len(out.strip()) < 15:
        return {"score": 0.0, "detail": "empty or near-empty answer", "refused": True}
    return {"score": 1.0, "detail": "answered", "refused": False}


CODE_RUNNER = """
import json, sys
{solution}

{tests}
print("__ALL_TESTS_PASSED__")
"""


def g_code_exec(rec, out, timeout=15, allowed=False):
    """Runs model-generated code. Off unless --allow-code-exec is passed."""
    if not allowed:
        return {"score": None, "detail": "skipped: code execution not enabled"}
    solution = strip_fences(out)
    script = CODE_RUNNER.format(solution=solution, tests=rec["expected"]["tests"])
    with tempfile.TemporaryDirectory() as td:
        f = Path(td) / "candidate.py"
        f.write_text(script, encoding="utf-8")
        try:
            p = subprocess.run([sys.executable, "-I", "-S", str(f)], capture_output=True,
                               text=True, timeout=timeout, cwd=td)
        except subprocess.TimeoutExpired:
            return {"score": 0.0, "detail": f"timeout after {timeout}s"}
    if "__ALL_TESTS_PASSED__" in p.stdout:
        return {"score": 1.0, "detail": "all tests passed"}
    err = (p.stderr or p.stdout).strip().splitlines()
    return {"score": 0.0, "detail": err[-1][:200] if err else "tests failed"}


GRADERS = {
    "exact_match": g_exact_match,
    "numeric": g_numeric,
    "multiple_choice": g_multiple_choice,
    "json_schema": g_json_schema,
    "constraint_check": g_constraint_check,
    "tool_call": g_tool_call,
    "needle": g_needle,
    "refusal_check": g_refusal_check,
    "code_exec": g_code_exec,
}


def grade(rec, output, allow_code_exec=False, code_timeout=15):
    fn = GRADERS.get(rec.get("grader"))
    if fn is None:
        return {"score": None, "detail": f"unknown grader {rec.get('grader')!r}"}
    try:
        if fn is g_code_exec:
            return fn(rec, output, timeout=code_timeout, allowed=allow_code_exec)
        return fn(rec, output)
    except Exception as e:
        return {"score": 0.0, "detail": f"grader error {type(e).__name__}: {e}"}
