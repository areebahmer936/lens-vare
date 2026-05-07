"""
LensVare — Audit Intelligence Backend
======================================
FastAPI service exposing the Excel → SQLite → Text-to-SQL → Prose pipeline
as a REST API consumed by the frontend.

Endpoints
---------
POST /upload            — Upload Excel file; returns session_id + schema summary
POST /ask               — Natural language query; returns prose + sql + row_count
GET  /sessions/{id}/schema  — Return the schema text for a session
DELETE /sessions/{id}   — Drop a session from memory
GET  /providers         — List available providers and their models
GET  /health            — Liveness probe for Docker/hosting health checks
"""

import re
import json
import uuid
import sqlite3
import traceback
from io import BytesIO
from typing import Optional

import pandas as pd
from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
from openai import OpenAI


# ──────────────────────────────────────────────────────────────────────────────
# App setup
# ──────────────────────────────────────────────────────────────────────────────

app = FastAPI(
    title="LensVare Audit Intelligence",
    description="Natural language querying over audit Excel data — zero hallucination.",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # tighten in production to your domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve the frontend static files at /
app.mount("/app", StaticFiles(directory="frontend", html=True), name="frontend")


@app.get("/")
def root():
    return FileResponse("frontend/index.html")


# ──────────────────────────────────────────────────────────────────────────────
# In-memory session store
# Each session holds: conn (SQLite), schema (str), history (list), df_cols (list)
# ──────────────────────────────────────────────────────────────────────────────

sessions: dict[str, dict] = {}   # session_id → {conn, schema, history, col_names}

SESSION_TTL_REQUESTS = 200        # auto-expire after N queries (simple guard)


# ──────────────────────────────────────────────────────────────────────────────
# Providers catalogue
# ──────────────────────────────────────────────────────────────────────────────

PROVIDERS: dict[str, dict] = {
    "Kilo": {
        "base_url": "https://api.kilo.ai/api/gateway",
        "models": [
            "qwen3.6-Plus",
            "qwen3.6-Flash",
            "qwen3.6-Open-Source",
            "qwen3.5-Plus",
            "qwen3.5-Flash",
            "qwen3.5-Open-Source",
            "Kimi k2.6",
            "Kimi k2.5",
        ],
        "key_url": "https://kilo.ai/dashboard",
        "note": "Kilo AI gateway with Qwen 3.5/3.6 and Kimi K2 series models.",
    },
    "DashScope (Qwen)": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "models": [
            "qwen3-235b-a22b",
            "qwen3-30b-a3b",
            "qwen-max",
            "qwen-plus",
            "qwen-turbo",
            "qwen2.5-72b-instruct",
        ],
        "key_url": "https://dashscope.console.aliyun.com/apiKey",
        "note": "Qwen3-235B-A22B is the flagship model with extended thinking.",
    },
    "Moonshot (Kimi)": {
        "base_url": "https://api.moonshot.cn/v1",
        "models": [
            "kimi-k2",
            "kimi-k2-instruct",
            "moonshot-v1-128k",
            "moonshot-v1-32k",
            "moonshot-v1-8k",
        ],
        "key_url": "https://platform.moonshot.cn/console/api-keys",
        "note": "Kimi K2 and Moonshot models. moonshot-v1-128k for very large datasets.",
    },
    "OpenRouter (all models)": {
        "base_url": "https://openrouter.ai/api/v1",
        "models": [
            "qwen/qwen3.6-plus",
            "qwen/qwen3.6-flash",
            "qwen/qwen3.5-plus",
            "qwen/qwen3.5-flash",
            "moonshotai/kimi-k2",
            "moonshotai/kimi-k2-instruct",
            "qwen/qwen3-235b-a22b",
            "qwen/qwen3-32b",
            "meta-llama/llama-3.3-70b-instruct",
            "deepseek/deepseek-r1",
            "google/gemma-4-26b-a4b-it:free",
            "google/gemma-4-31b-it:free",
            "openai/gpt-oss-120b:free",
            "z-ai/glm-4.5-air:free",
            "qwen/qwen3-coder:free",
        ],
        "key_url": "https://openrouter.ai/keys",
        "note": "One key for every model. Best for comparing providers.",
    },
    "Together.ai": {
        "base_url": "https://api.together.xyz/v1",
        "models": [
            "Qwen/Qwen3-235B-A22B",
            "Qwen/Qwen2.5-72B-Instruct-Turbo",
            "meta-llama/Meta-Llama-3.1-70B-Instruct-Turbo",
            "deepseek-ai/DeepSeek-R1",
        ],
        "key_url": "https://api.together.ai/settings/api-keys",
        "note": "Low-latency inference for open-weight models.",
    },
}


# ──────────────────────────────────────────────────────────────────────────────
# Data helpers (identical logic to app.py, now as pure functions)
# ──────────────────────────────────────────────────────────────────────────────

def _safe_col(name: str) -> str:
    s = re.sub(r"[^\w]", "_", name.strip())
    s = re.sub(r"_+", "_", s).strip("_")
    return ("col_" + s) if s and s[0].isdigit() else (s or "unnamed")


def load_excel_bytes(raw_bytes: bytes, filename: str) -> tuple[pd.DataFrame, sqlite3.Connection, dict]:
    engine = "xlrd" if filename.endswith(".xls") else "openpyxl"
    raw = pd.read_excel(BytesIO(raw_bytes), engine=engine)

    originals = raw.columns.tolist()
    seen: dict[str, int] = {}
    clean: list[str] = []
    for c in originals:
        base = _safe_col(str(c))
        if base in seen:
            seen[base] += 1
            clean.append(f"{base}_{seen[base]}")
        else:
            seen[base] = 0
            clean.append(base)

    raw.columns = clean
    col_map = dict(zip(originals, clean))

    conn = sqlite3.connect(":memory:", check_same_thread=False)
    raw.to_sql("audit_data", conn, if_exists="replace", index=False)
    return raw, conn, col_map


def build_schema(df: pd.DataFrame) -> str:
    lines = [
        "=== SQLITE DATABASE SCHEMA ===",
        "TABLE NAME : audit_data",
        f"TOTAL ROWS : {len(df):,}",
        "",
        "COLUMNS  (use these EXACT names in every SQL query)",
        "─" * 60,
    ]

    for col in df.columns:
        s = df[col]
        non_null = int(s.notna().sum())
        lines.append(f"\n{col}")

        if pd.api.types.is_numeric_dtype(s.dtype):
            lines.append("  type     : numeric")
            if non_null:
                lines.append(f"  range    : {s.min()} → {s.max()}")
                lines.append(f"  mean     : {s.mean():.2f}")
            lines.append(f"  non-null : {non_null:,} / {len(df):,}")

        elif pd.api.types.is_datetime64_any_dtype(s.dtype):
            lines.append("  type     : datetime")
            if non_null:
                lines.append(f"  range    : {s.min().date()} → {s.max().date()}")

        else:
            uniq = s.dropna().unique()
            n = len(uniq)
            lines.append(f"  type     : text  ({n} distinct values)")
            if n <= 40:
                vals = " | ".join(sorted(str(v) for v in uniq))
                lines.append(f"  values   : {vals}")
            else:
                top = s.dropna().value_counts().head(8).index.tolist()
                lines.append(f"  top-8    : {' | '.join(str(v) for v in top)}")
            lines.append(f"  non-null : {non_null:,} / {len(df):,}")

    lines += [
        "",
        "─" * 60,
        "SQL NOTES",
        "  • Partial text  : column LIKE '%value%'",
        "  • Case-insensitive : LOWER(column) LIKE LOWER('%value%')",
        "  • NULL check    : column IS NULL / IS NOT NULL",
        "  • Aggregation   : COUNT(*), SUM(col), AVG(col), GROUP BY col",
        "  • Sorting       : ORDER BY col DESC/ASC  LIMIT n",
        "=== END SCHEMA ===",
    ]
    return "\n".join(lines)


# ──────────────────────────────────────────────────────────────────────────────
# LLM pipeline
# ──────────────────────────────────────────────────────────────────────────────

def _strip_think(text: str) -> str:
    return re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()


def _extract_sql(raw: str) -> str:
    s = re.sub(r"```[a-z]*\n?", "", raw).replace("```", "").strip()
    m = re.search(r"(SELECT\b.+)", s, re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else s.strip()


def gen_sql(client: OpenAI, model: str, schema: str, question: str, history: list) -> str:
    system = f"""You are a precise SQLite query generator for an audit management system.

{schema}

STRICT OUTPUT RULES:
1. Return ONLY the raw SQL SELECT statement — nothing else.
2. No markdown, no explanation, no comments.
3. Use the exact column names listed in the schema.
4. For text matching: use LIKE or exact equality depending on context.
5. For "why" / "what causes" questions: SELECT all relevant columns including
   any root cause, department, risk rating, and description columns.
6. Never use INSERT, UPDATE, DELETE, DROP, CREATE, or any mutating statement.
7. If the question cannot be answered with SQL, write:
   SELECT 'Cannot answer this with SQL' AS message;"""

    msgs = [{"role": "system", "content": system}]
    for h in history[-3:]:
        msgs.append({"role": "user",      "content": h["q"]})
        msgs.append({"role": "assistant", "content": h["sql"]})
    msgs.append({"role": "user", "content": question})

    resp = client.chat.completions.create(
        model=model, messages=msgs, temperature=0.0, max_tokens=600,
    )
    return _extract_sql(_strip_think(resp.choices[0].message.content))


def run_sql(conn: sqlite3.Connection, sql: str) -> tuple[list[dict], str | None]:
    try:
        cur = conn.cursor()
        cur.execute(sql)
        if not cur.description:
            return [], None
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()], None
    except Exception as exc:
        return [], str(exc)


def gen_response(client: OpenAI, model: str, question: str, sql: str, results: list[dict]) -> str:
    n = len(results)
    if n == 0:
        data_str = "The query returned zero records."
    elif n <= 60:
        data_str = json.dumps(results, default=str, indent=2)
    else:
        data_str = (
            f"Total matching records: {n}\n\n"
            "First 50 records (representative sample):\n"
            + json.dumps(results[:50], default=str, indent=2)
            + f"\n\n[... {n - 50} additional records with similar structure]"
        )

    system = """You answer questions about audit data directly and concisely.

RULES:
• Use only facts from the query results — never invent numbers, names, or causes.
• For simple lookups (counts, names, values): give a direct one-sentence answer.
• For analytical questions (trends, why, patterns): give a brief analysis — 1 to 2 short paragraphs max.
• Never write generic boilerplate or pad the response.
• If zero records matched, say so and state what filter was applied."""

    user = (
        f"Question: {question}\n\n"
        f"SQL executed:\n{sql}\n\n"
        f"Data ({n} records):\n{data_str}\n\n"
        "Answer the question using only the data above. Be direct and brief."
    )

    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system},
            {"role": "user",   "content": user},
        ],
        temperature=0.3,
        max_tokens=900,
    )
    return _strip_think(resp.choices[0].message.content)


# ──────────────────────────────────────────────────────────────────────────────
# Request / Response models
# ──────────────────────────────────────────────────────────────────────────────

class AskRequest(BaseModel):
    session_id: str
    question: str
    provider: str
    model: str
    api_key: str


class AskResponse(BaseModel):
    answer: str
    sql: str
    row_count: int
    error: Optional[str] = None


class UploadResponse(BaseModel):
    session_id: str
    row_count: int
    col_count: int
    columns: list[str]
    schema_preview: str        # first 600 chars of the schema for display


# ──────────────────────────────────────────────────────────────────────────────
# API Routes
# ──────────────────────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return {"status": "ok", "sessions_active": len(sessions)}


@app.get("/providers")
def list_providers():
    return {
        name: {
            "models": data["models"],
            "key_url": data["key_url"],
            "note": data["note"],
        }
        for name, data in PROVIDERS.items()
    }


@app.post("/upload", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)):
    """
    Accept an Excel file upload.
    Parse it into SQLite, build schema, return a session_id.
    The session_id is used for all subsequent /ask calls.
    """
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Only .xlsx and .xls files are accepted.")

    content = await file.read()
    if len(content) > 50 * 1024 * 1024:  # 50 MB cap
        raise HTTPException(status_code=413, detail="File too large. Maximum 50 MB.")

    try:
        df, conn, col_map = load_excel_bytes(content, file.filename)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"Could not parse file: {exc}")

    schema = build_schema(df)
    session_id = str(uuid.uuid4())

    sessions[session_id] = {
        "conn": conn,
        "schema": schema,
        "history": [],
        "col_names": list(df.columns),
        "request_count": 0,
    }

    return UploadResponse(
        session_id=session_id,
        row_count=len(df),
        col_count=len(df.columns),
        columns=list(df.columns),
        schema_preview=schema[:800],
    )


@app.get("/sessions/{session_id}/schema")
def get_schema(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found or expired.")
    return {"schema": sessions[session_id]["schema"]}


@app.delete("/sessions/{session_id}")
def delete_session(session_id: str):
    if session_id in sessions:
        sessions[session_id]["conn"].close()
        del sessions[session_id]
    return {"deleted": True}


@app.post("/ask", response_model=AskResponse)
def ask(req: AskRequest):
    """
    Main query endpoint.

    Flow:
      1. Look up SQLite connection + schema for session_id
      2. Generate SQL using the LLM (temperature=0)
      3. Execute SQL against the in-memory SQLite database (real data, no hallucination)
      4. If SQL fails, retry once with the error message fed back
      5. Convert result rows into professional prose (temperature=0.3)
      6. Return prose + sql + row_count to frontend
    """
    if req.session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found. Please re-upload your file.")

    sess = sessions[req.session_id]

    # Guard against dangling large sessions
    sess["request_count"] += 1
    if sess["request_count"] > SESSION_TTL_REQUESTS:
        sess["history"] = sess["history"][-10:]  # keep only recent history

    # Validate provider
    if req.provider not in PROVIDERS:
        raise HTTPException(status_code=400, detail=f"Unknown provider: {req.provider}")

    prov = PROVIDERS[req.provider]
    try:
        client = OpenAI(api_key=req.api_key, base_url=prov["base_url"])
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not initialise client: {exc}")

    # ── Step 1: Generate SQL ──────────────────────────────────────────────────
    try:
        sql = gen_sql(client, req.model, sess["schema"], req.question, sess["history"])
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"LLM error generating SQL: {exc}")

    # ── Step 2: Execute SQL ───────────────────────────────────────────────────
    rows, err = run_sql(sess["conn"], sql)

    # Auto-retry once if the SQL had a syntax error
    if err:
        retry_prompt = (
            f"The SQL you generated caused this error:\n{err}\n\n"
            f"Faulty SQL:\n{sql}\n\n"
            f"Original question: {req.question}\n\n"
            "Return ONLY the corrected SQL SELECT statement."
        )
        try:
            sql = gen_sql(client, req.model, sess["schema"], retry_prompt, [])
            rows, err = run_sql(sess["conn"], sql)
        except Exception:
            pass

    if err:
        return AskResponse(
            answer=f"I was unable to retrieve an answer. The database query failed: {err}",
            sql=sql,
            row_count=0,
            error=err,
        )

    # ── Step 3: Generate prose response ──────────────────────────────────────
    try:
        answer = gen_response(client, req.model, req.question, sql, rows)
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"LLM error generating response: {exc}")

    # Update history for follow-up questions
    sess["history"].append({"q": req.question, "sql": sql})

    return AskResponse(
        answer=answer,
        sql=sql,
        row_count=len(rows),
    )


@app.post("/ask/stream")
def ask_stream(req: AskRequest):
    """
    Streaming version of /ask.  Emits SSE events:
      {type: "meta",  sql: "...", row_count: N}  — after SQL executes
      {type: "token", text: "..."}               — one prose chunk
      {type: "done"}                             — stream complete
      {type: "error", message: "..."}            — on failure
    """
    if req.session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found. Please re-upload your file.")

    sess = sessions[req.session_id]
    sess["request_count"] += 1
    if sess["request_count"] > SESSION_TTL_REQUESTS:
        sess["history"] = sess["history"][-10:]

    if req.provider not in PROVIDERS:
        raise HTTPException(status_code=400, detail=f"Unknown provider: {req.provider}")

    prov = PROVIDERS[req.provider]
    try:
        client = OpenAI(api_key=req.api_key, base_url=prov["base_url"])
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not initialise client: {exc}")

    def event_stream():
        # ── Detect @visualize flag ────────────────────────────────
        is_visualize = '@visualize' in req.question.lower()
        clean_q = re.sub(r'@visualize\s*', '', req.question, flags=re.IGNORECASE).strip()

        # ── Phase 1: Generate SQL (blocking, fast) ───────────────
        try:
            sql = gen_sql(client, req.model, sess["schema"], clean_q, sess["history"])
        except Exception as exc:
            yield f"data: {json.dumps({'type': 'error', 'message': str(exc)})}\n\n"
            return

        # ── Phase 2: Execute SQL ─────────────────────────────────
        rows, err = run_sql(sess["conn"], sql)
        if err:
            retry_prompt = (
                f"The SQL you generated caused this error:\n{err}\n\n"
                f"Faulty SQL:\n{sql}\n\n"
                f"Original question: {clean_q}\n\n"
                "Return ONLY the corrected SQL SELECT statement."
            )
            try:
                sql = gen_sql(client, req.model, sess["schema"], retry_prompt, [])
                rows, err = run_sql(sess["conn"], sql)
            except Exception:
                pass

        if err:
            yield f"data: {json.dumps({'type': 'meta', 'sql': sql, 'row_count': 0})}\n\n"
            yield f"data: {json.dumps({'type': 'error', 'message': err})}\n\n"
            return

        # Signal frontend: SQL ready, prose streaming starts now
        yield f"data: {json.dumps({'type': 'meta', 'sql': sql, 'row_count': len(rows)})}\n\n"

        # ── Phase 3: Chart or Prose ───────────────────────────────
        n = len(rows)
        if n == 0:
            data_str = "The query returned zero records."
        elif n <= 60:
            data_str = json.dumps(rows, default=str, indent=2)
        else:
            data_str = (
                f"Total matching records: {n}\n\n"
                "First 50 records (representative sample):\n"
                + json.dumps(rows[:50], default=str, indent=2)
                + f"\n\n[... {n - 50} additional records with similar structure]"
            )

        if is_visualize:
            # ── Chart generation (non-streaming) ─────────────────
            chart_system = (
                "You are a Chart.js v4 configuration generator.\n"
                "Output ONLY a valid Chart.js config JSON object — "
                "no prose, no markdown code fences, no explanation.\n\n"
                "RULES:\n"
                "- Use the actual data values from the query results.\n"
                "- Infer chart type from the request (pie, bar, line, doughnut, radar); default to bar.\n"
                "- For labels: use the first categorical/string column in the results.\n"
                "- For data values: use the first numeric column in the results.\n"
                "- backgroundColor array: [\"#4f6ef7\",\"#26c281\",\"#f5a623\",\"#e05252\","
                "\"#9b59b6\",\"#1abc9c\",\"#e67e22\",\"#3498db\"]\n"
                "- Set all text/tick/legend colors to \"#e8eaf6\" (dark theme).\n"
                "- For bar/line charts include scales with ticks.color and grid.color \"#2e3349\".\n"
                "- Output must be parseable by JSON.parse() with no surrounding text."
            )
            chart_user = (
                f"User request: {clean_q}\n\n"
                f"SQL executed:\n{sql}\n\n"
                f"Data ({n} records):\n{data_str}\n\n"
                "Output the Chart.js v4 config JSON object:"
            )
            try:
                resp = client.chat.completions.create(
                    model=req.model,
                    messages=[
                        {"role": "system", "content": chart_system},
                        {"role": "user",   "content": chart_user},
                    ],
                    temperature=0.1,
                    max_tokens=1500,
                )
                raw = _strip_think(resp.choices[0].message.content)
                raw = re.sub(r'^```(?:json)?\s*', '', raw.strip())
                raw = re.sub(r'\s*```$', '', raw.strip())
                config = json.loads(raw)
                yield f"data: {json.dumps({'type': 'chart', 'config': config})}\n\n"
            except json.JSONDecodeError:
                yield f"data: {json.dumps({'type': 'error', 'message': 'Could not generate chart — try a more specific @visualize request.'})}\n\n"
                return
            except Exception as exc:
                yield f"data: {json.dumps({'type': 'error', 'message': str(exc)})}\n\n"
                return

        else:
            # ── Prose streaming ───────────────────────────────────
            system = """You answer questions about audit data directly and concisely.

RULES:
• Use only facts from the query results — never invent numbers, names, or causes.
• For simple lookups (counts, names, values): give a direct one-sentence answer.
• For analytical questions (trends, why, patterns): give a brief analysis — 1 to 2 short paragraphs max.
• Never write generic boilerplate or pad the response.
• If zero records matched, say so and state what filter was applied."""

            user = (
                f"Question: {clean_q}\n\n"
                f"SQL executed:\n{sql}\n\n"
                f"Data ({n} records):\n{data_str}\n\n"
                "Answer the question using only the data above. Be direct and brief."
            )

            in_think = False
            think_buf = ""
            try:
                stream = client.chat.completions.create(
                    model=req.model,
                    messages=[
                        {"role": "system", "content": system},
                        {"role": "user",   "content": user},
                    ],
                    temperature=0.3,
                    max_tokens=900,
                    stream=True,
                )
                for chunk in stream:
                    delta = (chunk.choices[0].delta.content or "") if chunk.choices else ""
                    if not delta:
                        continue

                    if in_think:
                        think_buf += delta
                        if "</think>" in think_buf:
                            in_think = False
                            after = think_buf.split("</think>", 1)[1]
                            think_buf = ""
                            if after:
                                yield f"data: {json.dumps({'type': 'token', 'text': after})}\n\n"
                    else:
                        if "<think>" in delta:
                            parts = delta.split("<think>", 1)
                            before, rest = parts[0], parts[1]
                            in_think = True
                            think_buf = rest
                            if before:
                                yield f"data: {json.dumps({'type': 'token', 'text': before})}\n\n"
                        else:
                            yield f"data: {json.dumps({'type': 'token', 'text': delta})}\n\n"

            except Exception as exc:
                yield f"data: {json.dumps({'type': 'error', 'message': str(exc)})}\n\n"
                return

        sess["history"].append({"q": clean_q, "sql": sql})
        yield f"data: {json.dumps({'type': 'done'})}\n\n"

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )
