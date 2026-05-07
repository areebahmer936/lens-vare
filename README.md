# LensVare · Audit Intelligence

A production-ready web application that lets anyone — including non-technical executives — upload an Excel audit file and ask questions about it in plain English.

No Python knowledge required. No command-line. Just open a browser.

---

## How it works — The retrieval & query pipeline

This is the core question most people ask: *"How does the demo actually answer questions without making things up?"*

The answer is a strict **three-step pipeline** that keeps the LLM away from the raw data and forces it to query a real database instead.

```
┌───────────────────────────────────────────────────────────┐
│                   THREE-STEP PIPELINE                     │
│                                                           │
│  1. NL → SQL          2. Execute SQL        3. SQL Results → Prose │
│                                                           │
│  User question    →   Run against       →   LLM sees ONLY │
│  + schema context     real SQLite DB        actual rows   │
│                                             → prose answer │
│  (LLM, T=0.0)         (Python stdlib)      (LLM, T=0.3)  │
└───────────────────────────────────────────────────────────┘
```

### Step 1 — Natural Language → SQL  (temperature = 0)

The first LLM call receives:
- The auto-generated schema (column names, types, value ranges)
- The last 3 previous question→SQL pairs (conversation memory)
- The user's question

It produces a single `SELECT` statement — nothing else. Temperature is set to **0** so the output is deterministic and consistent.

### Step 2 — Execute SQL against real data

The generated SQL is run directly against an **in-memory SQLite database** that holds the exact rows from the uploaded Excel file. This is pure Python `sqlite3` — no network calls, no external DB.

If the SQL fails (syntax error, wrong column name), the error message is fed back to the LLM automatically for **one retry**. This handles common mistakes like wrong capitalisation of column names.

### Step 3 — Real rows → Prose answer  (temperature = 0.3)

The second LLM call sees:
- The original question
- The SQL that was executed
- The **actual rows returned** from SQLite (up to 200 rows)

It is instructed with a strict system prompt:

> *"Answer ONLY from the provided query results. Do not infer, extrapolate, or invent any figures not present in the data."*

Because the LLM only ever sees **real data that was actually returned by the database query**, it cannot hallucinate numbers or invent records that don't exist.

---

## Schema auto-generation

When a file is uploaded, LensVare automatically profiles every column:

| Column type | What is captured |
|---|---|
| **Numeric** | min, max, mean |
| **Categorical (≤40 unique values)** | ALL unique values listed |
| **Categorical (>40 unique values)** | Top 8 most frequent values |
| **Date/text** | Sample values |

This rich schema is injected into the SQL-generation prompt, so the LLM knows exactly which values are valid for WHERE clauses. For example, if a `Status` column contains `["Open", "Closed", "In Progress"]`, the LLM knows to use those exact strings.

---

## Architecture

```
Browser (CEO's laptop)
        │
        │  HTTP/REST  ─── POST /upload ───────────────────────────┐
        │               POST /ask                                  │
        │               GET /providers                             ▼
        │               GET /sessions/{id}/schema          ┌──────────────┐
        │               DELETE /sessions/{id}              │   FastAPI    │
        │                                                  │  (Python)    │
        │◄────────────────────────────────────────────────►│              │
                                                           │  SQLite      │
                                                           │  (in-memory) │
                                                           │              │
                                                           │  OpenAI SDK  │
                                                           │  (any LLM)   │
                                                           └──────────────┘
```

- **Frontend**: Pure HTML + CSS + vanilla JavaScript — no build step, no Node.js
- **Backend**: FastAPI (Python) serves both the REST API and the static frontend files
- **Storage**: SQLite per upload session, held in server memory (no disk writes)
- **LLM**: OpenAI-compatible API — works with DashScope, Moonshot, OpenRouter, Together.ai

---

## Supported LLM Providers

| Provider | Good for | Get API key |
|---|---|---|
| **DashScope (Qwen)** | Qwen3, Qwen-Max | [dashscope.aliyun.com](https://dashscope.aliyun.com) |
| **Moonshot (Kimi)** | Long-context Chinese/English | [platform.moonshot.cn](https://platform.moonshot.cn) |
| **OpenRouter** | Access 100+ models with one key | [openrouter.ai](https://openrouter.ai) |
| **Together.ai** | Fast open-source models | [together.ai](https://www.together.ai) |

All providers are accessed via the OpenAI Python SDK using their OpenAI-compatible endpoints.

> **Qwen3 / DeepSeek-R1 note**: These models emit `<think>…</think>` reasoning tokens before the answer. LensVare strips these automatically so only the clean answer reaches the user.

---

## Deployment

### Requirements

- Docker + Docker Compose

### Run

```bash
git clone <this-repo>
cd LensVare
docker compose up --build
```

Open [http://localhost:8000](http://localhost:8000) in your browser.

That's it. No Python installation, no environment setup on the client machine.

### Stop

```bash
docker compose down
```

---

## File structure

```
LensVare/
├── backend/
│   └── main.py          # FastAPI app — all server logic
├── frontend/
│   └── index.html       # Single-page UI — pure HTML/CSS/JS
├── app.py               # Original Streamlit prototype (reference only)
├── requirements.txt     # Python dependencies
├── Dockerfile           # Single-stage container build
├── docker-compose.yml   # One-command deployment
└── README.md
```

---

## Supported file formats

| Format | Extension | Notes |
|---|---|---|
| Excel (xlsx) | `.xlsx` | Requires openpyxl |
| Excel (xls)  | `.xls`  | Requires xlrd |

Maximum file size: **50 MB**. For larger files, consider splitting by date range or department before upload.

---

## Security notes

- API keys are entered in the browser and sent to the backend only to make LLM calls. They are **not stored on disk**.
- Sessions and uploaded data live in server memory and are lost when the container restarts.
- For production use, add HTTPS (e.g. via an nginx reverse proxy or a cloud load balancer) and restrict CORS origins in `backend/main.py`.

---

## Local development (without Docker)

```bash
pip install -r requirements.txt
uvicorn backend.main:app --reload --port 8000
```

Open [http://localhost:8000](http://localhost:8000).
