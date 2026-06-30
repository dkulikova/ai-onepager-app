from __future__ import annotations

import os
import sqlite3
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

BASE = Path(__file__).parent
DB_PATH = BASE / "briefing_demo.db"
NOTES_DIR = BASE / "notes"

st.set_page_config(page_title="AI Business Briefing Assistant", layout="wide")


# -----------------------------
# Data access layer
# -----------------------------
@st.cache_resource
def get_connection() -> sqlite3.Connection:
    """Return a cached SQLite connection for the local demo database."""
    return sqlite3.connect(DB_PATH, check_same_thread=False)


@st.cache_data
def load_table(table_name: str) -> pd.DataFrame:
    conn = get_connection()
    return pd.read_sql_query(f"SELECT * FROM {table_name}", conn)


def read_markdown_notes() -> Dict[str, str]:
    """Read local Obsidian/Markdown-style notes from the notes folder."""
    NOTES_DIR.mkdir(exist_ok=True)
    notes: Dict[str, str] = {}
    for file in NOTES_DIR.glob("*.md"):
        notes[file.name] = file.read_text(encoding="utf-8", errors="ignore")
    return notes


def retrieve_relevant_notes(query: str, notes: Dict[str, str], max_notes: int = 3) -> List[str]:
    """Simple keyword retrieval for demo purposes.

    Production version could replace this with embeddings/vector search.
    """
    query_terms = {term.lower().strip() for term in query.split() if len(term) > 2}
    scored = []
    for name, text in notes.items():
        haystack = f"{name}\n{text}".lower()
        score = sum(1 for term in query_terms if term in haystack)
        if score > 0:
            scored.append((score, name, text))
    scored.sort(reverse=True, key=lambda x: x[0])
    return [f"Source note: {name}\n{text}" for _, name, text in scored[:max_notes]]


# -----------------------------
# Prompt / workflow layer
# -----------------------------
def build_source_pack(
    selected_company: pd.Series,
    selected_project: pd.Series,
    extra_notes: str,
    retrieved_notes: List[str],
) -> str:
    company_context = "\n".join([f"- {col}: {selected_company[col]}" for col in selected_company.index])
    project_context = "\n".join([f"- {col}: {selected_project[col]}" for col in selected_project.index])
    note_context = "\n\n".join(retrieved_notes) if retrieved_notes else "No relevant local notes found."

    return f"""
STRUCTURED COMPANY DATA
{company_context}

PROJECT / USE CASE DATA
{project_context}

LOCAL MARKDOWN / OBSIDIAN NOTES
{note_context}

USER ADDED NOTES
{extra_notes or 'No additional notes provided.'}
""".strip()


def build_prompt(source_pack: str, audience: str, tone: str, brief_type: str) -> str:
    return f"""
You are helping draft a one-page business brief for a non-technical business audience.

Brief type: {brief_type}
Audience: {audience}
Tone: {tone}

Use ONLY the source pack below. Do not invent facts. If something is uncertain, say it is uncertain.

SOURCE PACK:
{source_pack}

Create a clear one-page brief with these sections:
1. Headline
2. Executive summary, 3-4 bullets
3. Key facts from the evidence
4. Why this matters
5. Recommended next steps
6. Things to verify / limitations

Style rules:
- Use plain English.
- Be concise but specific.
- Avoid hype.
- Separate evidence from interpretation.
- Make the output useful for a senior business stakeholder.
""".strip()


def build_review_prompt(draft: str, source_pack: str) -> str:
    return f"""
Review the draft below against the source pack.

Return:
1. Claims that appear unsupported or need checking
2. Important missing information
3. Suggestions to make the brief clearer for a non-technical audience

SOURCE PACK:
{source_pack}

DRAFT:
{draft}
""".strip()


# -----------------------------
# LLM provider layer
# -----------------------------
def get_secret(name: str) -> Optional[str]:
    # Try Streamlit secrets first, then environment variables.
    try:
        value = st.secrets.get(name)
        if value:
            return str(value)
    except Exception:
        pass
    return os.getenv(name)


def call_openai(prompt: str, model: str = "gpt-4.1-mini") -> str:
    from openai import OpenAI

    client = OpenAI(api_key=get_secret("OPENAI_API_KEY"))
    response = client.responses.create(model=model, input=prompt)
    return response.output_text


def call_anthropic(prompt: str, model: str = "claude-3-5-sonnet-latest") -> str:
    import anthropic

    client = anthropic.Anthropic(api_key=get_secret("ANTHROPIC_API_KEY"))
    message = client.messages.create(
        model=model,
        max_tokens=1600,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(block.text for block in message.content if getattr(block, "type", None) == "text")


def demo_fallback(source_pack: str) -> str:
    return f"""
# AI Business Brief: Demo Output

## Executive summary
- This is a prototype output generated without an API key.
- The workflow combines structured company/project data with unstructured Markdown notes.
- In a live version, an LLM would use the source pack to draft a tailored one-page brief.
- The final output should always be reviewed by a human before being shared.

## Key facts from the evidence
The tool has retrieved the following source context:

```text
{source_pack[:1800]}
```

## Why this matters
The visible output is a one-pager, but the underlying workflow demonstrates a broader AI pattern: connecting to data, retrieving relevant context, generating a structured draft, and flagging uncertainty.

## Recommended next steps
- Add a real API key for OpenAI or Anthropic.
- Replace keyword note search with embeddings/vector search if the note library grows.
- Add source citations and stricter validation before using the output externally.

## Things to verify / limitations
- This fallback output is not model-generated.
- Any production version would need data access controls, auditability and human approval.
""".strip()


def run_llm(prompt: str, provider: str) -> str:
    if provider == "OpenAI" and get_secret("OPENAI_API_KEY"):
        return call_openai(prompt)
    if provider == "Anthropic" and get_secret("ANTHROPIC_API_KEY"):
        return call_anthropic(prompt)
    return demo_fallback(prompt)


# -----------------------------
# Streamlit UI
# -----------------------------
st.title("AI Business Briefing Assistant")
st.caption("Prototype: structured data + Markdown notes + LLM workflow → reviewable one-page brief")

if not DB_PATH.exists():
    st.error("Demo database not found. Run `python setup_demo_data.py` first.")
    st.stop()

companies = load_table("companies")
projects = load_table("projects")
notes = read_markdown_notes()

with st.sidebar:
    st.header("Demo controls")
    provider = st.selectbox("LLM provider", ["Demo fallback", "OpenAI", "Anthropic"])
    brief_type = st.selectbox(
        "Brief type",
        ["Company opportunity brief", "Client meeting prep", "Project summary", "Executive decision note"],
    )
    audience = st.selectbox(
        "Audience",
        ["Senior business stakeholder", "Client-facing team", "Internal project team", "Non-technical executive"],
    )
    tone = st.selectbox(
        "Tone",
        ["Clear and professional", "Senior analyst style", "Concise and punchy", "Client-ready"],
    )

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Select business context")
    company_name = st.selectbox("Company", companies["company_name"].tolist())
    project_name = st.selectbox("Project/use case", projects["project_name"].tolist())

    selected_company = companies.loc[companies["company_name"] == company_name].iloc[0]
    selected_project = projects.loc[projects["project_name"] == project_name].iloc[0]

    with st.expander("View database record"):
        st.dataframe(pd.DataFrame(selected_company).rename(columns={0: "value"}))

with col2:
    st.subheader("2. Add extra context")
    extra_notes = st.text_area(
        "Paste messy notes, meeting notes, or instructions",
        height=180,
        placeholder="Example: Focus on why this company could be relevant for a client meeting next week...",
    )
    uploaded_file = st.file_uploader("Optional: upload a .txt or .md file", type=["txt", "md"])
    if uploaded_file is not None:
        extra_notes += "\n\nUPLOADED FILE CONTENT:\n" + uploaded_file.read().decode("utf-8", errors="ignore")

query = f"{company_name} {project_name} {extra_notes}"
retrieved_notes = retrieve_relevant_notes(query, notes)
source_pack = build_source_pack(selected_company, selected_project, extra_notes, retrieved_notes)

st.subheader("3. Generate output")
st.write("This runs a simple chain: retrieve context → draft one-pager → review draft.")

if st.button("Generate one-pager", type="primary"):
    with st.spinner("Building source pack and drafting brief..."):
        draft_prompt = build_prompt(source_pack, audience, tone, brief_type)
        draft = run_llm(draft_prompt, provider if provider != "Demo fallback" else "Demo fallback")

    with st.spinner("Running review step..."):
        review_prompt = build_review_prompt(draft, source_pack)
        review = run_llm(review_prompt, provider if provider != "Demo fallback" else "Demo fallback")

    st.session_state["draft"] = draft
    st.session_state["review"] = review
    st.session_state["source_pack"] = source_pack

if "draft" in st.session_state:
    tab1, tab2, tab3 = st.tabs(["One-pager", "AI review", "Source pack"])
    with tab1:
        st.markdown(st.session_state["draft"])
        st.download_button(
            "Download one-pager as Markdown",
            data=st.session_state["draft"],
            file_name="ai_one_pager.md",
            mime="text/markdown",
        )
    with tab2:
        st.markdown(st.session_state["review"])
    with tab3:
        st.code(st.session_state["source_pack"], language="text")

st.divider()
st.caption(
    "Interview talking point: the prototype uses a lightweight local database and Markdown notes. "
    "A production version could connect to CRM, SharePoint, Snowflake, SQL Server or MCP-style connectors, with permissions and audit controls."
)
