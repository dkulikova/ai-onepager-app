# AI Business Briefing Assistant

A lightweight Streamlit prototype for an interview demo.

The demo shows how an AI-assisted one-pager tool can combine:

- structured data from a local SQLite database
- unstructured Markdown / Obsidian-style notes
- user-provided messy notes
- an LLM generation step
- a review / quality-check step

The visible output is a one-page business brief, but the more interesting point is the workflow behind it: source retrieval, structured prompting, drafting, and human review.

## 1. Setup

```bash
cd onepager_ai_app
python -m venv .venv
source .venv/bin/activate  # Mac/Linux
# .venv\Scripts\activate   # Windows
pip install -r requirements.txt
python setup_demo_data.py
streamlit run app.py
```

## 2. API keys

The app works without an API key in `Demo fallback` mode.

To use a real model, copy the example secrets file:

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Then add either:

```toml
OPENAI_API_KEY = "your-key-here"
```

or:

```toml
ANTHROPIC_API_KEY = "your-key-here"
```

Do not commit real keys to GitHub.

## 3. What to demo

Suggested demo flow:

1. Select a company and project from the dropdowns.
2. Add messy notes in the text box.
3. Generate the one-pager.
4. Show the generated brief.
5. Show the AI review tab.
6. Show the source pack tab to explain grounding.

## 4. How to explain the architecture

Plain-English version:

> The tool acts like a fast first-draft analyst. It pulls together structured data, notes and user context, then drafts a one-page brief. A second review step checks for unsupported claims, missing information and uncertainty.

Technical version:

```text
Streamlit UI
   ↓
SQLite database + Markdown notes
   ↓
Simple retrieval step
   ↓
LLM prompt chain
   ↓
One-pager draft
   ↓
AI review / human-in-the-loop check
```

## 5. Future improvements

For a stronger production version, you could add:

- embeddings / vector search instead of keyword note retrieval
- database connectors to CRM, SharePoint, Snowflake or SQL Server
- MCP-style connectors for approved external systems
- source citations for every claim
- user authentication and permission-aware retrieval
- audit logs and version history
- export to PowerPoint, Word or PDF
- human approval workflow before client use

## 6. Interview positioning

Suggested framing:

> The one-pager is just the visible output. The real system is an AI workflow that retrieves context, structures information, creates a draft and flags where human judgement is still required.

Limitations to mention:

- output quality depends on input quality
- LLMs can sound confident even when wrong
- the prototype uses simple keyword retrieval rather than semantic search
- source grounding and permission controls would be critical in production
- human review remains essential
