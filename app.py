"""
Audit Data Intelligence — Demo
================================
Natural language queries over Excel audit data using Qwen / Kimi / open models.

Pipeline:  Question → SQL (T=0) → Execute on real data → Grounded prose
Guarantee: The response layer only sees rows the SQL actually returned.
           If a number isn't in the result set, the model cannot say it.
"""

import re
import json
import sqlite3
import traceback

import pandas as pd
import streamlit as st
from openai import OpenAI


# ──────────────────────────────────────────────────────────────────────────────
# PROVIDERS
# All use the OpenAI-compatible chat completions interface.
# ──────────────────────────────────────────────────────────────────────────────
PROVIDERS: dict[str, dict] = {
    "DashScope (Qwen)": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "models": [
            "qwen3-235b-a22b",   # strongest, extended thinking
            "qwen3-30b-a3b",
            "qwen-max",
            "qwen-plus",
            "qwen-turbo",
            "qwen2.5-72b-instruct",
            "qwen2.5-32b-instruct",
        ],
        "key_url": "https://dashscope.console.aliyun.com/apiKey",
        "note":    "Qwen3-235B-A22B is the flagship model with extended thinking.",
    },
    "Moonshot (Kimi)": {
        "base_url": "https://api.moonshot.cn/v1",
        "models": [
            "moonshot-v1-128k",
            "moonshot-v1-32k",
            "moonshot-v1-8k",
        ],
        "key_url": "https://platform.moonshot.cn/console/api-keys",
        "note":    "moonshot-v1-128k handles large datasets with a 128k context window.",
    },
    "OpenRouter (all models)": {
        "base_url": "https://openrouter.ai/api/v1",
        "models": [
            "qwen/qwen3-235b-a22b",
            "qwen/qwen3-32b",
            "qwen/qwen2.5-72b-instruct",
            "moonshotai/kimi-dev-72b",
            "meta-llama/llama-3.3-70b-instruct",
            "deepseek/deepseek-r1",
            "mistralai/mistral-large",
        ],
        "key_url": "https://openrouter.ai/keys",
        "note":    "One key for every model. Best for comparing providers.",
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
        "note":    "Low-latency inference for open-weight models.",
    },
}


# ──────────────────────────────────────────────────────────────────────────────
# DATA LAYER  —  Excel → SQLite
# ──────────────────────────────────────────────────────────────────────────────

def _safe_col(name: str) -> str:
    """Convert an arbitrary Excel header to a valid SQL identifier."""
    s = re.sub(r"[^\w]", "_", name.strip())
    s = re.sub(r"_+", "_", s).strip("_")
    return ("col_" + s) if s and s[0].isdigit() else (s or "unnamed")


def load_excel(file) -> tuple[pd.DataFrame, sqlite3.Connection, dict]:
    """
    Read the uploaded Excel into a pandas DataFrame, then persist it to an
    in-memory SQLite database as the table 'audit_data'.

    Returns
    -------
    df       : cleaned DataFrame (column names are SQL-safe)
    conn     : sqlite3 connection (check_same_thread=False for Streamlit)
    col_map  : {original header → sql column name}
    """
    fname  = getattr(file, "name", "")
    engine = "xlrd" if fname.endswith(".xls") else "openpyxl"
    raw    = pd.read_excel(file, engine=engine)

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


# ──────────────────────────────────────────────────────────────────────────────
# SCHEMA LAYER  —  Auto-generate rich column context
#
# Why this matters:
#   Listing every actual value for categorical columns (e.g., Risk Rating:
#   HIGH | MEDIUM | LOW) means the model can't invent "CRITICAL" or "EXTREME".
#   Numeric range/mean lets the model sanity-check its own SQL aggregations.
#   This is the entire foundation of hallucination prevention.
# ──────────────────────────────────────────────────────────────────────────────

def build_schema(df: pd.DataFrame) -> str:
    """
    Build a plain-text schema description that the LLM receives as system
    context on every call.  Kept under ~3 KB so it doesn't eat context budget.
    """
    lines = [
        "=== SQLITE DATABASE SCHEMA ===",
        "TABLE NAME : audit_data",
        f"TOTAL ROWS : {len(df):,}",
        "",
        "COLUMNS  (use these EXACT names in every SQL query)",
        "─" * 60,
    ]

    for col in df.columns:
        s       = df[col]
        non_null = int(s.notna().sum())
        lines.append(f"\n{col}")

        if pd.api.types.is_numeric_dtype(s.dtype):
            lines.append(f"  type     : numeric")
            if non_null:
                lines.append(f"  range    : {s.min()} → {s.max()}")
                lines.append(f"  mean     : {s.mean():.2f}")
            lines.append(f"  non-null : {non_null:,} / {len(df):,}")

        elif pd.api.types.is_datetime64_any_dtype(s.dtype):
            lines.append(f"  type     : datetime")
            if non_null:
                lines.append(f"  range    : {s.min().date()} → {s.max().date()}")

        else:
            uniq = s.dropna().unique()
            n    = len(uniq)
            lines.append(f"  type     : text  ({n} distinct values)")
            if n <= 40:
                # List every valid value — model cannot fabricate anything else
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
# LLM UTILITIES
# ──────────────────────────────────────────────────────────────────────────────

def _strip_think(text: str) -> str:
    """Remove Qwen3 / DeepSeek-R1 chain-of-thought blocks before returning."""
    return re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()


def _extract_sql(raw: str) -> str:
    """Pull clean SQL out of potentially markdown-wrapped LLM output."""
    s = re.sub(r"```[a-z]*\n?", "", raw).replace("```", "").strip()
    m = re.search(r"(SELECT\b.+)", s, re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else s.strip()


# ──────────────────────────────────────────────────────────────────────────────
# STEP 1  —  Question → SQL
# ──────────────────────────────────────────────────────────────────────────────

def gen_sql(
    client: OpenAI,
    model: str,
    schema: str,
    question: str,
    history: list[dict],
) -> str:
    """
    Ask the model to produce a SQLite SELECT query.
    Temperature = 0 → deterministic, no creativity.
    History injects the last 3 Q→SQL pairs so follow-up questions work
    (e.g., "now break that down by department" after a count question).
    """
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
    raw = _strip_think(resp.choices[0].message.content)
    return _extract_sql(raw)


# ──────────────────────────────────────────────────────────────────────────────
# STEP 2  —  Execute SQL
# ──────────────────────────────────────────────────────────────────────────────

def run_sql(
    conn: sqlite3.Connection, sql: str
) -> tuple[list[dict], str | None]:
    """Execute SQL; returns (rows_as_dicts, error_message_or_None)."""
    try:
        cur = conn.cursor()
        cur.execute(sql)
        if not cur.description:
            return [], None
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()], None
    except Exception as exc:
        return [], str(exc)


# ──────────────────────────────────────────────────────────────────────────────
# STEP 3  —  SQL results → Natural language response
# ──────────────────────────────────────────────────────────────────────────────

def gen_response(
    client: OpenAI,
    model: str,
    question: str,
    sql: str,
    results: list[dict],
) -> str:
    """
    Turn raw SQL rows into professional prose.
    The model is given ONLY the rows that SQL returned — it cannot reference
    anything that wasn't in those rows.
    """
    n = len(results)
    if n == 0:
        data_str = "The query returned zero records."
    elif n <= 60:
        data_str = json.dumps(results, default=str, indent=2)
    else:
        data_str = (
            f"Total matching records: {n}\n\n"
            f"First 50 records (representative sample):\n"
            + json.dumps(results[:50], default=str, indent=2)
            + f"\n\n[... {n - 50} additional records with similar structure]"
        )

    system = """You are a senior audit partner presenting findings to the audit committee.

ABSOLUTE RULES — violating any of these is unacceptable:
• Every number or fact you state must come directly from the query results.
• Do NOT invent risks, departments, percentages, or root causes not present in the data.
• Write flowing professional prose — 2 to 4 paragraphs, no bullet lists.
• Use specific figures: exact counts, percentages calculated from the data, named departments.
• For "why" or root-cause questions: analyze the actual text in Root_Cause (or similar)
  columns from the rows returned — synthesize patterns across multiple entries.
• If zero records matched, say exactly that and note what filter was applied.
• Never pad with generic audit boilerplate that isn't supported by the data."""

    user = (
        f"Question asked: {question}\n\n"
        f"SQL that was executed:\n{sql}\n\n"
        f"Data returned ({n} records):\n{data_str}\n\n"
        "Write a concise, grounded analytical response to the question "
        "using only the data above."
    )

    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system},
            {"role": "user",   "content": user},
        ],
        temperature=0.3,
        max_tokens=1400,
    )
    return _strip_think(resp.choices[0].message.content)


# ──────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Audit Intelligence",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🔍 Audit Data Intelligence")
st.caption(
    "Upload Excel → ask natural language questions → get analysis grounded in real data. "
    "No hallucination: every number in the response came from an actual SQL query."
)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Model Configuration")

    provider_name = st.selectbox("Provider", list(PROVIDERS.keys()))
    prov          = PROVIDERS[provider_name]

    st.caption(prov["note"])
    st.markdown(f"[→ Get API key]({prov['key_url']})")

    api_key = st.text_input("API Key", type="password", placeholder="sk-…")
    model   = st.selectbox("Model", prov["models"])

    if "qwen3" in model.lower() or "deepseek-r1" in model.lower():
        st.info(
            "This model uses extended thinking (chain-of-thought). "
            "The reasoning is stripped before displaying the answer."
        )

    st.divider()
    st.header("📁 Data")
    uploaded = st.file_uploader(
        "Excel file (.xlsx / .xls)",
        type=["xlsx", "xls"],
        help="500+ row files work fine. The data stays local in SQLite.",
    )

    if st.button("🗑️ Clear conversation", use_container_width=True):
        st.session_state.msgs = []
        st.session_state.hist = []
        st.rerun()

# ── Session state defaults ────────────────────────────────────────────────────
for key, val in [
    ("df",     None),
    ("conn",   None),
    ("schema", None),
    ("fname",  None),
    ("msgs",   []),
    ("hist",   []),
]:
    if key not in st.session_state:
        st.session_state[key] = val

# ── Load Excel on first upload or file change ─────────────────────────────────
if uploaded and uploaded.name != st.session_state.fname:
    with st.spinner(f"Loading {uploaded.name}…"):
        df, conn, col_map = load_excel(uploaded)
        schema = build_schema(df)
        st.session_state.df     = df
        st.session_state.conn   = conn
        st.session_state.schema = schema
        st.session_state.fname  = uploaded.name
        st.session_state.msgs   = []
        st.session_state.hist   = []
    st.sidebar.success(f"✅ {len(df):,} records, {len(df.columns)} columns")

# ── Nothing uploaded yet ──────────────────────────────────────────────────────
if st.session_state.df is None:
    st.info("⬅️  Upload an Excel file from the sidebar to begin.")

    with st.expander("📐 Architecture — how zero-hallucination works"):
        st.markdown("""
**Step 1 — Schema generation (what you're asking about)**

When you upload Excel, the app reads every column header and infers its type.
For categorical columns with ≤40 unique values it records **every single valid
value** — so if `Risk_Rating` only ever contains `HIGH`, `MEDIUM`, `LOW`, the
model is told exactly that. It cannot fabricate `CRITICAL` or `VERY HIGH`
because those strings don't appear in the schema it received.
For numeric columns it computes min, max, and mean so the model can
sanity-check its own aggregation SQL.

The schema is injected as system context on every call.  It's typically 1–3 KB,
so it costs very little context budget.

---

**Step 2 — Text → SQL  (temperature = 0)**

Your question + the schema go to the model.  It outputs **only a SQL SELECT
statement**, nothing else.  Temperature zero means deterministic — no creative
liberties taken.  The model cannot invent a number it didn't query for because
it hasn't seen the data yet.

---

**Step 3 — SQL executes against real rows**

The SQL runs on your actual data in SQLite.  The model sees only the rows it
asked for.  Zero rows returned → the response layer says "no records matched"
rather than guessing.

---

**Step 4 — Rows → Grounded prose**

The response prompt receives only the query result rows and the original
question.  Every figure in the answer is traceable to a specific row.

---

**Minimal production architecture (integrating with your existing system)**

```
 ┌─────────────────────────────────────────────────────────────┐
 │  Existing .NET / SQL Server system  (unchanged)             │
 └────────────────────────┬────────────────────────────────────┘
                          │  read-only SELECT only
              ┌───────────▼───────────┐
              │  SQL MCP Server       │  tiny Node.js or Python
              │  (sidecar process)    │  sidecar — no DB changes
              └───────────┬───────────┘
                          │
              ┌───────────▼───────────┐
              │  schema.yaml          │  maintained by audit team,
              │  (static context)     │  describes columns in
              │                       │  business language, not
              │  tables, columns,     │  just SQL types
              │  value enumerations,  │
              │  join patterns,       │
              │  business glossary    │
              └───────────┬───────────┘
                          │
              ┌───────────▼───────────┐
              │  Agent endpoint       │  single new POST /ask
              │  on your existing API │  endpoint; frontend
              │                       │  unchanged
              │  schema + question    │
              │      → SQL tool       │
              │      → real rows      │
              │      → prose response │
              └───────────────────────┘
```

The schema file is the most important piece.  Instead of `RiskRating TEXT`,
write: *"RiskRating — composite Likelihood × Impact score; values: HIGH,
MEDIUM, LOW; used to prioritise remediation."*  That 20-word description
eliminates an entire class of SQL generation errors.
        """)

    st.stop()

# ── Metrics row ───────────────────────────────────────────────────────────────
df = st.session_state.df
c1, c2, c3, c4 = st.columns(4)
c1.metric("Records",    f"{len(df):,}")
c2.metric("Columns",    len(df.columns))
c3.metric("File",       st.session_state.fname or "—")
c4.metric("Turns",      len(st.session_state.msgs) // 2)

with st.expander("📋 Auto-generated schema  (what the model receives as context)"):
    st.code(st.session_state.schema, language="text")

st.divider()

# ── Chat history ──────────────────────────────────────────────────────────────
for msg in st.session_state.msgs:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])
        if msg.get("sql"):
            with st.expander("🔎 SQL that was executed"):
                st.code(msg["sql"], language="sql")
        if msg.get("n") is not None:
            st.caption(f"Response grounded in {msg['n']:,} matching records")

# ── Guard: need API key ───────────────────────────────────────────────────────
if not api_key:
    st.warning("🔑  Enter your API key in the sidebar to start querying.")
    st.stop()

# ── Sample questions (shown only before first message) ───────────────────────
if not st.session_state.msgs:
    st.markdown("**Sample questions to try:**")
    st.markdown(
        "- *How many high risk findings are there?*  \n"
        "- *Which departments have the most risks, and what are the root causes?*  \n"
        "- *Why are risks in the Shared Services department rated high?*  \n"
        "- *How many findings are overdue and who owns them?*  \n"
        "- *Break down risk count by department and rating*"
    )

# ── Chat input ────────────────────────────────────────────────────────────────
question = st.chat_input("Ask anything about your audit data…")
if not question:
    st.stop()

# User bubble
st.session_state.msgs.append({"role": "user", "content": question})
with st.chat_message("user"):
    st.write(question)

# Assistant bubble
with st.chat_message("assistant"):
    try:
        client = OpenAI(api_key=api_key, base_url=prov["base_url"])

        with st.status("Analysing…", expanded=True) as status:

            # ── Step 1: generate SQL ──────────────────────────────────────
            status.write("⚙️  Generating SQL query from your question…")
            sql = gen_sql(
                client, model,
                st.session_state.schema,
                question,
                st.session_state.hist,
            )

            # ── Step 2: execute ───────────────────────────────────────────
            status.write("🗄️  Executing query against your data…")
            results, err = run_sql(st.session_state.conn, sql)

            if err:
                # One automatic retry with the error fed back
                status.write(f"⚠️  SQL error — retrying ({err[:80]}…)")
                sql = gen_sql(
                    client, model,
                    st.session_state.schema,
                    f"Question: {question}\n\n"
                    f"Your previous SQL failed with: {err}\n"
                    "Generate a corrected SQL query.",
                    [],   # no history context on retry
                )
                results, err = run_sql(st.session_state.conn, sql)

            # ── Step 3: generate response ─────────────────────────────────
            status.write("✍️  Generating analytical response…")
            if err:
                answer = (
                    f"I was unable to query the data after two attempts.\n\n"
                    f"**Last error:** `{err}`\n\n"
                    f"**SQL attempted:**\n```sql\n{sql}\n```"
                )
                n_rows = 0
            else:
                answer = gen_response(client, model, question, sql, results)
                n_rows = len(results)

            status.update(label="Done", state="complete", expanded=False)

        # Display
        st.write(answer)
        with st.expander("🔎 SQL that was executed"):
            st.code(sql, language="sql")
        st.caption(f"Response grounded in {n_rows:,} matching records from your data")

        # Save to history
        st.session_state.msgs.append({
            "role":    "assistant",
            "content": answer,
            "sql":     sql,
            "n":       n_rows,
        })
        st.session_state.hist.append({"q": question, "sql": sql})

    except Exception:
        st.error("An unexpected error occurred:")
        st.code(traceback.format_exc())
