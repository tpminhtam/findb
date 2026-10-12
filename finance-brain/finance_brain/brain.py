"""Ask the brain: retrieve (local or HydraDB) -> Claude answers with citations, lineage chains and impact lists."""
from __future__ import annotations

import os

import anthropic

MODEL = os.environ.get("FB_MODEL", "claude-opus-5-5")
EFFORT = os.environ.get("FB_EFFORT", "low")

SYSTEM = """You are finDB, the shared brain of a finance team, built from their spreadsheets.
Answer only from the CONTEXT (table cards, inputs, notes, cell comments, findings, lineage records, graph paths, analyst memories).
Never invent cells, numbers or owners. If the context does not cover the question, say what is missing in one sentence.

Format, strictly:
- Line 1: the direct answer in one short bold sentence. Say yes or no first when the question allows it.
- Then at most 4 bullets, one short line each, with the evidence that matters: cell addresses, row counts, definitions, owners.
- "Where does X come from": the bullets are the chain from the number back to its raw inputs, in order, naming the function at each step.
- "What breaks": the bullets are the affected outputs, most important first.
- End a bullet with one source number like [3]. Use no other brackets.
- No headings, no tables, no preamble, no closing summary. 80 words at most."""


def build_prompt(question: str, ctx: dict, persona: str | None) -> str:
    parts = []
    if ctx.get("memories"):
        parts.append("ANALYST MEMORIES (" + (persona or "analyst") + "):\n" + "\n".join(f"- {m['text']}" for m in ctx["memories"]))
    if ctx.get("context_md"):
        parts.append("CONTEXT (results are numbered; cite them as [n]):\n" + ctx["context_md"])
    else:
        parts.append("CONTEXT (numbered sources):\n" + "\n".join(f"[{i}] {c['text']}" for i, c in enumerate(ctx.get("chunks", []), start=1)))
    if ctx.get("paths"):
        parts.append("GRAPH PATHS (<- means 'derives from', -> means 'feeds'):\n" + "\n".join(ctx["paths"]))
    parts.append(f"QUESTION: {question}")
    return "\n\n".join(parts)


def retrieval_only_answer(ctx: dict) -> str:
    lines = ["Retrieved context (no generated answer):"]
    for c in ctx.get("chunks", [])[:6]:
        lines.append(f"- [{c['id']}] {c['text'][:300]}")
    for p in ctx.get("paths", [])[:8]:
        lines.append(f"- {p}")
    return "\n".join(lines)


def ask(question: str, store, persona: str | None = "analyst", graph: bool = True, k: int = 10,
        model: str | None = None, effort: str | None = None) -> dict:
    ctx = store.query(question, persona=persona, k=k, graph=graph)
    prompt = build_prompt(question, ctx, persona)
    out = {"question": question, "persona": persona, "graph": graph, "chunks": ctx.get("chunks", []), "paths": ctx.get("paths", []),
           "memories": ctx.get("memories", []), "model": model or MODEL, "store": ctx.get("store", "local")}
    client = anthropic.Anthropic()   # ANTHROPIC_API_KEY, or the `ant auth login` profile on this machine
    system = [{"type": "text", "text": SYSTEM, "cache_control": {"type": "ephemeral"}}]
    messages = [{"role": "user", "content": prompt}]
    try:
        try:
            resp = client.beta.messages.create(
                model=model or MODEL, max_tokens=2500, output_config={"effort": effort or EFFORT},
                betas=["server-side-fallback-2026-07-01"], fallbacks="default", system=system, messages=messages)
        except anthropic.BadRequestError:
            resp = client.messages.create(model=model or MODEL, max_tokens=2500, output_config={"effort": effort or EFFORT}, system=system, messages=messages)
    except (anthropic.AuthenticationError, anthropic.CredentialsError) as e:
        out.update(answer=retrieval_only_answer(ctx) + f"\n\n(Claude auth failed: {getattr(e, "message", e)}. Put a valid ANTHROPIC_API_KEY in .env.)", mode="retrieval-only")
        return out
    except anthropic.APIConnectionError as e:
        out.update(answer=retrieval_only_answer(ctx) + f"\n\n(Claude unreachable: {e}.)", mode="retrieval-only")
        return out
    if resp.stop_reason == "refusal":
        out.update(answer="The model declined this request.", mode="refusal")
        return out
    text = "".join(b.text for b in resp.content if b.type == "text")
    out.update(answer=text.strip(), mode="generated", usage={"input": resp.usage.input_tokens, "output": resp.usage.output_tokens,
               "cache_read": getattr(resp.usage, "cache_read_input_tokens", 0)}, model=resp.model)
    return out
