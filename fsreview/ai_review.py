"""Optional AI proofreading with Claude.

Runs only when the user turns it on and the 'anthropic' package and API
credentials are available. Claude reviews the narrative text (policies,
notes, reports) for grammar, clarity and disclosure wording, and returns
suggestions only: nothing it says is applied to the document, and it is told
never to propose changes to figures.
"""

import importlib.util
import json

from .model import Issue

MODEL = "claude-opus-5-5"

SYSTEM = (
    "You are a senior financial statements reviewer. You proofread narrative text from a set of financial "
    "statements prepared under {framework}{jurisdiction}. Report only real problems: spelling, grammar, "
    "unclear or ambiguous wording, inconsistent terminology, wrong standard names or numbers, and disclosure "
    "wording that does not fit the stated framework. Never suggest changing any figure, amount, date or "
    "note number; if one looks wrong, describe the concern instead. Quote the exact text you refer to."
)

SCHEMA = {
    "type": "object",
    "properties": {
        "findings": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "paragraph": {"type": "integer"},
                    "quote": {"type": "string"},
                    "problem": {"type": "string"},
                    "suggestion": {"type": "string"},
                    "severity": {"type": "string", "enum": ["error", "warning", "info"]},
                },
                "required": ["paragraph", "quote", "problem", "suggestion", "severity"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["findings"],
    "additionalProperties": False,
}


def available():
    return importlib.util.find_spec("anthropic") is not None


def ai_proofread(doc, settings, max_chars=60000):
    if not available():
        return [Issue("proofreading", "info", "AI review was requested but the 'anthropic' package is not "
                      "installed (pip install anthropic). Rule-based proofreading was still run.")]
    import anthropic

    paras = [(i, b) for i, b in enumerate(doc.blocks)
             if b.kind == "paragraph" and b.role in ("body", "heading", "note_heading") and len(b.text) > 40]
    batches, cur, size = [], [], 0
    for i, b in paras:
        if size + len(b.text) > max_chars and cur:
            batches.append(cur)
            cur, size = [], 0
        cur.append((i, b))
        size += len(b.text)
    if cur:
        batches.append(cur)
    client = anthropic.Anthropic()
    system = SYSTEM.format(framework=settings.get("framework") or "IFRS",
                           jurisdiction=f" in {settings['jurisdiction']}" if settings.get("jurisdiction") else "")
    issues = []
    for batch in batches:
        text = "\n\n".join(f"[{i}] {b.text}" for i, b in batch)
        try:
            resp = client.beta.messages.create(
                model=MODEL,
                max_tokens=16000,
                system=system,
                betas=["server-side-fallback-2026-07-01"],
                fallbacks="default",
                output_config={"effort": "medium",
                               "format": {"type": "json_schema", "schema": SCHEMA}},
                messages=[{"role": "user", "content": "Proofread these paragraphs. Each starts with its "
                                                      "number in brackets.\n\n" + text}],
            )
        except (anthropic.AuthenticationError, TypeError):
            # TypeError: the SDK found no credentials at all
            return issues + [Issue("proofreading", "warning", "AI review skipped: no valid Anthropic API key "
                                   "(set ANTHROPIC_API_KEY).")]
        except anthropic.APIConnectionError:
            return issues + [Issue("proofreading", "warning", "AI review skipped: could not reach the Anthropic "
                                   "API.")]
        except anthropic.APIStatusError as e:
            return issues + [Issue("proofreading", "warning", f"AI review stopped: API error {e.status_code}.")]
        if resp.stop_reason == "refusal":
            issues.append(Issue("proofreading", "info", "AI review declined part of the text; that part was "
                                "covered by the rule-based checks only."))
            continue
        out = next((b.text for b in resp.content if b.type == "text"), "")
        try:
            data = json.loads(out)
        except json.JSONDecodeError:
            continue
        by_index = dict(batch)
        for f in data.get("findings", []):
            b = by_index.get(f.get("paragraph"))
            issues.append(Issue("proofreading", f.get("severity", "info"),
                                f"AI review: {f['problem']} — “{f['quote'][:160]}”",
                                location=b.source if b else "", section=b.section if b else "",
                                suggestion=f.get("suggestion", "")))
    return issues
