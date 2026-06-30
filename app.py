from __future__ import annotations

import json
import os
import re
import sqlite3
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Any

import pandas as pd
import streamlit as st

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.util import Inches, Pt

BASE = Path(__file__).parent


def resolve_db_path() -> Path:
    """Support both foldered repos and flat GitHub uploads."""
    candidates = [
        BASE / "briefing_demo.db",
        BASE / "briefing_demo.db",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    # Default creation path if setup_demo_data is needed.
    return BASE / "briefing_demo.db"


def resolve_notes_dirs() -> List[Path]:
    """Support notes in /notes and notes uploaded at repo root."""
    candidates = [BASE / "notes", BASE]
    return [path for path in candidates if path.exists()]


DB_PATH = resolve_db_path()
NOTES_DIRS = resolve_notes_dirs()

st.set_page_config(page_title="AI Business Briefing Assistant", layout="wide")


# -----------------------------
# Data access layer
# -----------------------------
@st.cache_resource
def get_connection(db_path_str: str) -> sqlite3.Connection:
    """Return a cached SQLite connection for the local demo database."""
    return sqlite3.connect(db_path_str, check_same_thread=False)


@st.cache_data
def load_table(table_name: str, db_path_str: str) -> pd.DataFrame:
    conn = get_connection(db_path_str)
    return pd.read_sql_query(f"SELECT * FROM {table_name}", conn)


def read_markdown_notes() -> Dict[str, str]:
    """Read local Obsidian/Markdown-style notes.

    The app supports both the intended /notes folder and a flat GitHub upload where
    the .md files sit beside app.py. README.md is excluded.
    """
    notes: Dict[str, str] = {}
    for notes_dir in resolve_notes_dirs():
        for file in notes_dir.glob("*.md"):
            if file.name.lower() == "readme.md":
                continue
            if file.name in notes:
                continue
            notes[file.name] = file.read_text(encoding="utf-8", errors="ignore")
    return notes


def retrieve_relevant_notes(query: str, notes: Dict[str, str], max_notes: int = 3) -> List[str]:
    """Simple keyword retrieval for demo purposes.

    Production version could replace this with embeddings/vector search.
    """
    query_terms = {term.lower().strip() for term in re.split(r"\W+", query) if len(term) > 2}
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
You are helping draft a one-page company profile for a non-technical business audience.

Brief type: {brief_type}
Audience: {audience}
Tone: {tone}

Use ONLY the source pack below. Do not invent facts. If something is uncertain, say it is uncertain.

SOURCE PACK:
{source_pack}

Return ONLY valid JSON using exactly this schema:
{{
  "headline": "short title for the profile",
  "company_positioning": "1-2 sentences on the company's positioning / mission",
  "growth_direction": "1-2 sentences on likely growth direction / vision",
  "target_market": "1-2 sentences on target market / audience",
  "company_description": "2-3 sentences suitable for a left-side profile panel",
  "what_they_do": ["bullet 1", "bullet 2", "bullet 3"],
  "key_facts": ["bullet 1", "bullet 2", "bullet 3", "bullet 4"],
  "signals": ["bullet 1", "bullet 2", "bullet 3"],
  "risks": ["bullet 1", "bullet 2", "bullet 3"],
  "timeline": [
    {{"year": "2024", "text": "milestone or available evidence"}},
    {{"year": "2025", "text": "milestone or next step"}},
    {{"year": "2026", "text": "future opportunity / to verify"}}
  ],
  "next_steps": ["bullet 1", "bullet 2", "bullet 3"]
}}

Style rules:
- Use plain English.
- Be concise but specific.
- Avoid hype.
- Separate evidence from interpretation.
- Make the output useful for a senior business stakeholder.
- Keep each bullet short enough to fit on a PowerPoint one-pager.
""".strip()


def build_review_prompt(profile_json: str, source_pack: str) -> str:
    return f"""
Review the profile below against the source pack.

Return:
1. Claims that appear unsupported or need checking
2. Important missing information
3. Suggestions to make the brief clearer for a non-technical audience

SOURCE PACK:
{source_pack}

PROFILE JSON:
{profile_json}
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
        max_tokens=1800,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(block.text for block in message.content if getattr(block, "type", None) == "text")


def demo_fallback(source_pack: str) -> str:
    """Deterministic fallback used when no API key is configured."""
    return """
{
  "headline": "AI-Generated Company Profile: Demo Output",
  "company_positioning": "The company appears relevant based on its sector, operating footprint and available growth signals. This is a first-draft profile generated from structured data and local notes.",
  "growth_direction": "Available signals point to commercial scaling and potential need for more client-facing capacity. Any property requirement should be treated as inferred rather than confirmed.",
  "target_market": "The target audience for this brief is a senior business stakeholder who needs a quick view of why the company may be relevant.",
  "company_description": "This profile combines structured database fields with unstructured Markdown notes. It is designed as a business-ready first draft rather than a final recommendation.",
  "what_they_do": ["Uses available company and project records", "Combines structured data with analyst-style notes", "Turns messy inputs into a PowerPoint-ready brief"],
  "key_facts": ["Database and notes were retrieved successfully", "The visible output is a single-slide company profile", "The workflow can be extended to CRM, SharePoint or SQL systems", "Human review is required before external use"],
  "signals": ["Growth signal is inferred from the provided source pack", "The app can include user-added notes and uploaded text files", "A review step flags uncertainty and gaps"],
  "risks": ["Fallback mode is not using a live LLM", "The output depends on source quality and data freshness", "Claims should be checked before client use"],
  "timeline": [
    {"year": "2024", "text": "Structured demo data available"},
    {"year": "2025", "text": "Prototype one-pager workflow"},
    {"year": "2026", "text": "Potential connector expansion"}
  ],
  "next_steps": ["Add an OpenAI or Anthropic API key", "Replace demo data with approved internal sources", "Add source citations and approval workflow"]
}
""".strip()


def run_llm(prompt: str, provider: str) -> str:
    if provider == "OpenAI" and get_secret("OPENAI_API_KEY"):
        return call_openai(prompt)
    if provider == "Anthropic" and get_secret("ANTHROPIC_API_KEY"):
        return call_anthropic(prompt)
    return demo_fallback(prompt)


# -----------------------------
# Profile parsing / formatting
# -----------------------------
def extract_json(text: str) -> Optional[Dict[str, Any]]:
    """Extract JSON from an LLM response, including responses wrapped in code fences."""
    try:
        return json.loads(text)
    except Exception:
        pass

    match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, flags=re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except Exception:
            pass

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        try:
            return json.loads(text[start : end + 1])
        except Exception:
            pass
    return None


def as_list(value: Any, max_items: int = 5) -> List[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()][:max_items]
    lines = [line.strip(" -•\t") for line in str(value).split("\n") if line.strip()]
    return lines[:max_items]


def default_profile_sections(selected_company: pd.Series, selected_project: pd.Series) -> Dict[str, Any]:
    company = selected_company.to_dict()
    project = selected_project.to_dict()
    return {
        "headline": f"One Page Company Profile: {company.get('company_name', 'Company')}",
        "company_positioning": f"{company.get('company_name', 'The company')} operates in {company.get('sector', 'its sector')} and is headquartered in {company.get('hq_location', 'an unknown location')}.",
        "growth_direction": str(company.get("expansion_signal", "Growth direction needs to be verified.")),
        "target_market": str(project.get("target_sector", "Target market to be confirmed.")),
        "company_description": f"{company.get('company_name', 'The company')} is a {company.get('sector', 'company')} business with approximately {company.get('employee_count', 'unknown')} employees. Current office context: {company.get('current_office_location', 'not available')}.",
        "what_they_do": [
            f"Operates in {company.get('sector', 'sector not available')}",
            str(company.get("recent_news", "Recent activity to verify")),
            str(project.get("brief_description", "Use case to verify")),
        ],
        "key_facts": [
            f"HQ: {company.get('hq_location', 'Not available')}",
            f"Employees: {company.get('employee_count', 'Not available')}",
            f"Current office: {company.get('current_office_location', 'Not available')}",
            f"Relevance score: {company.get('relevance_score', 'Not available')}",
        ],
        "signals": [
            str(company.get("recent_news", "Recent news not available")),
            str(company.get("expansion_signal", "Expansion signal to verify")),
            "Relevance is inferred from available source material, not a confirmed requirement.",
        ],
        "risks": [
            "No confirmed real estate requirement unless stated in the source material.",
            "Validate latest company activity and current office footprint.",
            "Treat the AI output as a first draft for human review.",
        ],
        "timeline": [
            {"year": "2024", "text": "Available company evidence"},
            {"year": "2025", "text": "Potential outreach / review"},
            {"year": "2026", "text": "Future opportunity to verify"},
        ],
        "next_steps": [
            "Review source evidence.",
            "Validate assumptions with the business team.",
            "Refine before sharing externally.",
        ],
    }


def normalise_profile_sections(raw_text: str, selected_company: pd.Series, selected_project: pd.Series) -> Dict[str, Any]:
    base = default_profile_sections(selected_company, selected_project)
    parsed = extract_json(raw_text)
    if not parsed:
        return base

    for key in base:
        if key in parsed and parsed[key]:
            base[key] = parsed[key]
    return base


def profile_to_markdown(profile: Dict[str, Any]) -> str:
    timeline_lines = []
    for item in as_list(profile.get("timeline"), max_items=5):
        if isinstance(item, dict):
            timeline_lines.append(f"- **{item.get('year', '')}:** {item.get('text', '')}")
        else:
            timeline_lines.append(f"- {item}")
    if not timeline_lines and isinstance(profile.get("timeline"), list):
        for item in profile.get("timeline", []):
            if isinstance(item, dict):
                timeline_lines.append(f"- **{item.get('year', '')}:** {item.get('text', '')}")

    def bullets(key: str, max_items: int = 5) -> str:
        return "\n".join([f"- {x}" for x in as_list(profile.get(key), max_items=max_items)])

    return f"""
# {profile.get('headline', 'One Page Company Profile')}

## Company positioning
{profile.get('company_positioning', '')}

## Growth direction
{profile.get('growth_direction', '')}

## Target market
{profile.get('target_market', '')}

## Company description
{profile.get('company_description', '')}

## What they do
{bullets('what_they_do')}

## Key facts
{bullets('key_facts')}

## Relevant signals / evidence
{bullets('signals')}

## Risks / things to verify
{bullets('risks')}

## Timeline / next steps
{"\n".join(timeline_lines)}

## Recommended next steps
{bullets('next_steps')}
""".strip()


# -----------------------------
# PowerPoint generation layer
# -----------------------------
def add_textbox(slide, x, y, w, h, text="", font_size=10, bold=False, color=None, align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]
    p.text = str(text)
    p.font.size = Pt(font_size)
    p.font.bold = bold
    if color:
        p.font.color.rgb = color
    if align:
        p.alignment = align
    return box


def add_bullet_box(slide, x, y, w, h, items, font_size=8.5, color=None, max_items=4):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    items = as_list(items, max_items=max_items)
    if not items:
        items = ["Not available / to verify."]
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = str(item)
        p.font.size = Pt(font_size)
        if color:
            p.font.color.rgb = color
        p.level = 0
        p.space_after = Pt(3)
    return box


def add_card(slide, x, y, w, h, title, body_items, icon_text="", max_items=4):
    navy = RGBColor(25, 52, 78)
    teal = RGBColor(79, 161, 190)
    dark = RGBColor(45, 55, 65)

    # Icon circle
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y + 0.05), Inches(0.46), Inches(0.46))
    circle.fill.solid()
    circle.fill.fore_color.rgb = teal
    circle.line.fill.background()
    add_textbox(slide, x + 0.08, y + 0.15, 0.30, 0.18, icon_text, font_size=8, bold=True, color=RGBColor(255, 255, 255), align=PP_ALIGN.CENTER)

    add_textbox(slide, x + 0.62, y, w - 0.62, 0.45, title, font_size=13, bold=True, color=navy)
    add_bullet_box(slide, x + 0.62, y + 0.55, w - 0.62, h - 0.55, body_items, font_size=8.5, color=dark, max_items=max_items)


def add_profile_pptx(profile: Dict[str, Any], company: pd.Series, project: pd.Series) -> BytesIO:
    """Create a single-slide, editable PowerPoint one-pager."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    navy = RGBColor(18, 47, 76)
    deep_blue = RGBColor(22, 63, 101)
    teal = RGBColor(79, 161, 190)
    light_blue = RGBColor(223, 237, 244)
    light_grey = RGBColor(242, 242, 242)
    mid_grey = RGBColor(220, 225, 230)
    dark = RGBColor(37, 49, 59)
    white = RGBColor(255, 255, 255)

    # Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = white
    bg.line.fill.background()

    # Faint top image-style panels recreated with shapes
    top_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.55))
    top_bg.fill.solid()
    top_bg.fill.fore_color.rgb = light_blue
    top_bg.line.fill.background()
    for x, alpha_width in [(5.7, 0.15), (6.35, 0.09), (7.1, 0.12), (7.8, 0.08), (8.45, 0.12), (9.4, 0.10), (10.0, 0.12), (10.9, 0.10), (11.7, 0.12)]:
        line = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(x), Inches(0), Inches(alpha_width), Inches(1.55))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(175, 205, 220)
        line.line.fill.background()

    add_textbox(slide, 0.45, 0.28, 4.9, 0.45, "ONE PAGE COMPANY PROFILE", font_size=23, bold=False, color=navy)

    # Top dark ribbon
    ribbon = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(2.45), Inches(1.05), Inches(10.55), Inches(1.05))
    ribbon.fill.solid()
    ribbon.fill.fore_color.rgb = deep_blue
    ribbon.line.fill.background()

    top_sections = [
        (3.65, "P", "Positioning", profile.get("company_positioning", "")),
        (6.75, "G", "Growth direction", profile.get("growth_direction", "")),
        (9.85, "T", "Target market", profile.get("target_market", "")),
    ]
    for x, icon, title, body in top_sections:
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(1.31), Inches(0.52), Inches(0.52))
        circ.fill.solid(); circ.fill.fore_color.rgb = teal; circ.line.fill.background()
        add_textbox(slide, x + 0.11, 1.43, 0.30, 0.18, icon, font_size=8.5, bold=True, color=white, align=PP_ALIGN.CENTER)
        add_textbox(slide, x + 0.65, 1.28, 2.0, 0.24, title, font_size=12, bold=True, color=white)
        add_textbox(slide, x + 0.65, 1.58, 2.25, 0.42, str(body), font_size=8.2, color=white)

    # Left information panel
    panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.48), Inches(2.45), Inches(2.85), Inches(4.55))
    panel.fill.solid(); panel.fill.fore_color.rgb = deep_blue; panel.line.fill.background()

    company_name = str(company.get("company_name", "Company"))
    facts = [
        ("Company Name", company_name),
        ("HQ Location", company.get("hq_location", "N/A")),
        ("Industry", company.get("sector", "N/A")),
        ("Employees", company.get("employee_count", "N/A")),
        ("Office", company.get("current_office_location", "N/A")),
        ("Score", company.get("relevance_score", "N/A")),
    ]
    y = 2.65
    for label, value in facts:
        add_textbox(slide, 0.68, y, 0.95, 0.26, label, font_size=7.8, bold=True, color=white)
        add_textbox(slide, 1.72, y, 1.35, 0.26, str(value), font_size=7.6, color=white, align=PP_ALIGN.RIGHT)
        divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.68), Inches(y + 0.32), Inches(2.35), Inches(0.01))
        divider.fill.solid(); divider.fill.fore_color.rgb = RGBColor(170, 190, 205); divider.line.fill.background()
        y += 0.45

    add_textbox(slide, 0.68, 5.55, 2.25, 0.3, "Company Description", font_size=12, bold=True, color=white)
    add_textbox(slide, 0.68, 5.95, 2.25, 0.82, profile.get("company_description", ""), font_size=8.3, color=white)

    # Main section cards
    add_card(slide, 3.65, 2.42, 2.75, 1.40, "What they do", profile.get("what_they_do"), "W", max_items=3)
    add_card(slide, 6.85, 2.42, 2.75, 1.40, "Key facts", profile.get("key_facts"), "K", max_items=4)
    add_card(slide, 10.05, 2.42, 2.8, 1.40, "Evidence signals", profile.get("signals"), "E", max_items=3)

    add_card(slide, 3.65, 4.18, 2.75, 1.25, "Risks / verify", profile.get("risks"), "R", max_items=3)
    add_card(slide, 6.85, 4.18, 2.75, 1.25, "Next steps", profile.get("next_steps"), "N", max_items=3)
    add_card(slide, 10.05, 4.18, 2.8, 1.25, "Project context", [
        f"Use case: {project.get('project_name', 'N/A')}",
        f"Audience: {project.get('client', 'N/A')}",
        f"Location: {project.get('location', 'N/A')}",
    ], "C", max_items=3)

    # Timeline block
    add_textbox(slide, 3.65, 5.88, 3.3, 0.3, "Company milestones / next steps", font_size=13, bold=True, color=navy)
    timeline_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.65), Inches(6.25), Inches(9.2), Inches(0.95))
    timeline_bg.fill.solid(); timeline_bg.fill.fore_color.rgb = light_grey; timeline_bg.line.color.rgb = mid_grey

    timeline = profile.get("timeline", [])
    if not isinstance(timeline, list) or not timeline:
        timeline = default_profile_sections(company, project).get("timeline", [])
    timeline = timeline[:5]
    x_positions = [4.15, 5.95, 7.75, 9.55, 11.35]
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.4), Inches(6.63), Inches(6.95), Inches(0.025))
    line.fill.solid(); line.fill.fore_color.rgb = teal; line.line.fill.background()

    for i, item in enumerate(timeline):
        if not isinstance(item, dict):
            item = {"year": str(2024 + i), "text": str(item)}
        x = x_positions[i]
        add_textbox(slide, x - 0.25, 6.33, 0.9, 0.25, item.get("year", str(2024 + i)), font_size=11, bold=True, color=navy, align=PP_ALIGN.CENTER)
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(6.57), Inches(0.16), Inches(0.16))
        dot.fill.solid(); dot.fill.fore_color.rgb = white; dot.line.color.rgb = teal; dot.line.width = Pt(2)
        add_textbox(slide, x - 0.45, 6.78, 1.05, 0.32, item.get("text", ""), font_size=6.8, color=dark, align=PP_ALIGN.CENTER)

    # Footer
    add_textbox(slide, 0.48, 7.18, 12.35, 0.16, "AI-generated first draft — requires human review before external use.", font_size=7.5, color=RGBColor(90, 90, 90), align=PP_ALIGN.RIGHT)

    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io


# -----------------------------
# Streamlit UI
# -----------------------------
st.title("AI Business Briefing Assistant")
st.caption("Prototype: structured data + Markdown notes + LLM workflow → editable PowerPoint one-page profile")

if not DB_PATH.exists():
    # Streamlit Cloud may not have the prebuilt SQLite file if it was not uploaded to GitHub.
    # In that case, create the demo database and notes automatically on startup.
    try:
        import setup_demo_data

        setup_demo_data.main()
        DB_PATH = resolve_db_path()
        st.info("Demo database was missing, so it has been created automatically.")
    except Exception as e:
        st.error(f"Demo database could not be created automatically: {e}")
        st.stop()

companies = load_table("companies", str(DB_PATH))
projects = load_table("projects", str(DB_PATH))
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
    st.divider()
    st.caption(f"Database path: {DB_PATH.relative_to(BASE) if DB_PATH.exists() else DB_PATH}")
    st.caption(f"Markdown notes found: {len(notes)}")

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
st.write("This runs a simple chain: retrieve context → draft structured profile → review draft → export editable PowerPoint.")

if st.button("Generate one-pager", type="primary"):
    with st.spinner("Building source pack and drafting structured profile..."):
        draft_prompt = build_prompt(source_pack, audience, tone, brief_type)
        raw_draft = run_llm(draft_prompt, provider if provider != "Demo fallback" else "Demo fallback")
        profile = normalise_profile_sections(raw_draft, selected_company, selected_project)
        markdown_draft = profile_to_markdown(profile)

    with st.spinner("Running review step..."):
        review_prompt = build_review_prompt(json.dumps(profile, indent=2), source_pack)
        review = run_llm(review_prompt, provider if provider != "Demo fallback" else "Demo fallback")

    pptx_file = add_profile_pptx(profile, selected_company, selected_project)

    st.session_state["raw_draft"] = raw_draft
    st.session_state["profile"] = profile
    st.session_state["draft"] = markdown_draft
    st.session_state["review"] = review
    st.session_state["source_pack"] = source_pack
    st.session_state["pptx_file"] = pptx_file.getvalue()
    st.session_state["pptx_name"] = f"{company_name.replace(' ', '_')}_one_page_company_profile.pptx"

if "profile" in st.session_state:
    tab1, tab2, tab3, tab4 = st.tabs(["PowerPoint output", "Text preview", "AI review", "Source pack"])
    with tab1:
        st.success("PowerPoint one-pager created. Download and open in PowerPoint to edit the slide.")
        st.download_button(
            "Download one-page company profile as PowerPoint",
            data=st.session_state["pptx_file"],
            file_name=st.session_state.get("pptx_name", "one_page_company_profile.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        with st.expander("View structured fields used in the slide"):
            st.json(st.session_state["profile"])
    with tab2:
        st.markdown(st.session_state["draft"])
        st.download_button(
            "Download text preview as Markdown",
            data=st.session_state["draft"],
            file_name="ai_one_pager_preview.md",
            mime="text/markdown",
        )
    with tab3:
        st.markdown(st.session_state["review"])
    with tab4:
        st.code(st.session_state["source_pack"], language="text")

st.divider()
st.caption(
    "Interview talking point: the visible output is an editable PowerPoint one-pager, but the underlying pattern is source retrieval, "
    "structured extraction, LLM drafting, review and human approval. A production version could connect to CRM, SharePoint, "
    "Snowflake, SQL Server or MCP-style connectors, with permissions and audit controls."
)
