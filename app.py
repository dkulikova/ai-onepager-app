from __future__ import annotations

import json
import os
import re
import sqlite3
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import quote_plus

import feedparser
import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Inches, Pt

BASE = Path(__file__).parent


def resolve_db_path() -> Path:
    candidates = [BASE / "data" / "briefing_demo.db", BASE / "briefing_demo.db"]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return BASE / "data" / "briefing_demo.db"


def resolve_notes_dirs() -> List[Path]:
    candidates = [BASE / "notes", BASE]
    return [path for path in candidates if path.exists()]


DB_PATH = resolve_db_path()
st.set_page_config(page_title="AI Business Briefing Assistant", layout="wide")


# -----------------------------
# Data access layer
# -----------------------------
@st.cache_resource
def get_connection(db_path_str: str) -> sqlite3.Connection:
    return sqlite3.connect(db_path_str, check_same_thread=False)


@st.cache_data
def safe_load_table(table_name: str, db_path_str: str) -> pd.DataFrame:
    conn = get_connection(db_path_str)
    try:
        return pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
    except Exception:
        return pd.DataFrame()


def read_markdown_notes() -> Dict[str, str]:
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
    query_terms = {term.lower().strip() for term in re.split(r"\W+", query) if len(term) > 2}
    scored = []
    for name, text in notes.items():
        haystack = f"{name}\n{text}".lower()
        score = sum(1 for term in query_terms if term in haystack)
        if score > 0:
            scored.append((score, name, text))
    scored.sort(reverse=True, key=lambda x: x[0])
    return [f"Source note: {name}\n{text}" for _, name, text in scored[:max_notes]]


def table_for_company(table_name: str, company: pd.Series) -> pd.DataFrame:
    df = safe_load_table(table_name, str(DB_PATH))
    if df.empty or "company_id" not in df.columns or "company_id" not in company.index:
        return pd.DataFrame()
    company_id = company.get("company_id")
    return df.loc[df["company_id"].astype(str) == str(company_id)].copy()


def company_intelligence(company: pd.Series) -> Dict[str, pd.DataFrame]:
    return {
        "leadership": table_for_company("leadership", company),
        "products": table_for_company("products", company),
        "financials": table_for_company("financials", company),
        "milestones": table_for_company("milestones", company),
    }


def format_df_records(title: str, df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return f"{title}: No internal records found."
    lines = [title]
    for _, row in df.iterrows():
        values = []
        for col, val in row.items():
            if col.endswith("_id") or col == "company_id":
                continue
            values.append(f"{col}: {val}")
        lines.append("- " + "; ".join(values))
    return "\n".join(lines)


@st.cache_data(ttl=1800, show_spinner=False)
def get_latest_company_articles(company_name: str, max_articles: int = 5) -> List[Dict[str, str]]:
    query = quote_plus(f'"{company_name}"')
    rss_url = f"https://news.google.com/rss/search?q={query}&hl=en-GB&gl=GB&ceid=GB:en"

    try:
        feed = feedparser.parse(rss_url)
    except Exception as exc:
        return [{"title": "Latest news retrieval failed", "published": "", "summary": f"The RSS feed could not be read: {exc}", "link": rss_url}]

    articles: List[Dict[str, str]] = []
    seen_titles = set()
    for entry in feed.entries:
        title = entry.get("title", "").strip()
        if not title or title.lower() in seen_titles:
            continue
        seen_titles.add(title.lower())
        articles.append({
            "title": title,
            "published": entry.get("published", ""),
            "summary": re.sub(r"<.*?>", "", entry.get("summary", "")).strip(),
            "link": entry.get("link", ""),
        })
        if len(articles) >= max_articles:
            break

    if not articles:
        return [{"title": "No recent articles retrieved", "published": "", "summary": "No RSS results were returned for this company.", "link": rss_url}]
    return articles


def format_articles_for_prompt(articles: List[Dict[str, str]]) -> str:
    if not articles:
        return "Latest news was excluded or no recent articles were retrieved."
    formatted = []
    for i, article in enumerate(articles, start=1):
        formatted.append(
            f"""
Article {i}
Title: {article.get('title', '')}
Published: {article.get('published', '')}
Summary: {article.get('summary', '')}
Link: {article.get('link', '')}
""".strip()
        )
    return "\n\n".join(formatted)


# -----------------------------
# Matching / project context
# -----------------------------
def find_company_match(company_query: str, companies_df: pd.DataFrame) -> tuple[pd.Series, bool, str]:
    query = (company_query or "").strip() or "Unknown company"
    names = companies_df["company_name"].astype(str).tolist() if not companies_df.empty else []

    lower_map = {name.lower(): name for name in names}
    if query.lower() in lower_map:
        matched_name = lower_map[query.lower()]
        row = companies_df.loc[companies_df["company_name"] == matched_name].iloc[0]
        return row, True, f"Exact internal database match found: {matched_name}."

    for name in names:
        if query.lower() in name.lower() or name.lower() in query.lower():
            row = companies_df.loc[companies_df["company_name"] == name].iloc[0]
            return row, True, f"Likely internal database match found: {name}."

    scored = [(SequenceMatcher(None, query.lower(), name.lower()).ratio(), name) for name in names]
    scored.sort(reverse=True)
    if scored and scored[0][0] >= 0.72:
        matched_name = scored[0][1]
        row = companies_df.loc[companies_df["company_name"] == matched_name].iloc[0]
        return row, True, f"Fuzzy internal database match found: {matched_name} ({scored[0][0]:.0%} similarity)."

    synthetic = {col: "Not available" for col in companies_df.columns} if not companies_df.empty else {}
    synthetic.update({
        "company_id": "external_query",
        "company_name": query,
        "sector": "Not available from internal database",
        "hq_location": "Not available from internal database",
        "founded_year": "Not available",
        "company_type": "Not available",
        "employee_count": "Not available",
        "website": "Not available",
        "current_office_location": "Not available",
        "mission": "No internal mission record found.",
        "short_description": "No internal company profile found. Use user context and optional latest news cautiously.",
        "recent_news": "No internal database record found.",
        "expansion_signal": "No internal expansion signal found. Requires external verification.",
        "target_market": "Not available",
        "differentiation": "Not available",
        "relevance_score": "N/A",
    })
    return pd.Series(synthetic), False, "No confident internal database match found. The one-pager will use user context and optional latest news."


def build_project_context(brief_type: str, extra_notes: str) -> pd.Series:
    descriptions = {
        "Investment one-pager": "Create a concise investment-style profile covering mission, leadership, products, funding/commercial signals, risks and next diligence steps.",
        "Company brief": "Create a concise company profile covering mission, leadership, what the company does, key facts, recent signals and business relevance.",
    }
    return pd.Series({
        "project_id": "user_defined",
        "project_name": brief_type,
        "client": "User-defined / interview demo",
        "location": "Not specified",
        "target_sector": "To be inferred from the company and context",
        "brief_description": descriptions.get(brief_type, "Create a concise business one-pager."),
        "user_context": extra_notes or "No additional context provided.",
    })


# -----------------------------
# Source pack / prompts
# -----------------------------
def build_source_pack(
    selected_company: pd.Series,
    selected_project: pd.Series,
    extra_notes: str,
    retrieved_notes: List[str],
    latest_articles: List[Dict[str, str]],
    intel: Dict[str, pd.DataFrame],
) -> str:
    company_context = "\n".join([f"- {col}: {selected_company[col]}" for col in selected_company.index])
    project_context = "\n".join([f"- {col}: {selected_project[col]}" for col in selected_project.index])
    note_context = "\n\n".join(retrieved_notes) if retrieved_notes else "No relevant local notes found."
    latest_news_context = format_articles_for_prompt(latest_articles)

    return f"""
STRUCTURED COMPANY DATA
{company_context}

EXECUTIVE LEADERSHIP DATA
{format_df_records('Leadership records', intel.get('leadership', pd.DataFrame()))}

PRODUCTS / SERVICES DATA
{format_df_records('Product records', intel.get('products', pd.DataFrame()))}

FUNDING / COMMERCIAL SIGNALS DATA
{format_df_records('Financial and commercial records', intel.get('financials', pd.DataFrame()))}

MILESTONES DATA
{format_df_records('Milestone records', intel.get('milestones', pd.DataFrame()))}

PROJECT / USE CASE DATA
{project_context}

LOCAL MARKDOWN / OBSIDIAN NOTES
{note_context}

LATEST NEWS / RECENT EXTERNAL SIGNALS
{latest_news_context}

USER ADDED NOTES
{extra_notes or 'No additional notes provided.'}
""".strip()


def build_prompt(source_pack: str, brief_type: str) -> str:
    """Main generation prompt for the OpenAI call.

    This is the core prompt-engineering layer. It tells the model what role to play,
    what evidence to trust, how to treat private-company financial information, and
    exactly which structured fields to return for the PowerPoint template.
    """
    return f"""
You are a senior business analyst creating a one-page company intelligence brief for a non-technical senior audience.

Brief type: {brief_type}

Your job is to turn the source pack into a concise, evidence-grounded PowerPoint one-pager.
You are not writing marketing copy. You are creating a practical business briefing that a human can review, edit and use.

Core rules:
- Use ONLY the information in the source pack.
- Do not invent facts, numbers, partnerships, products, leaders, dates or financials.
- If the source pack is incomplete, say what is missing or what needs verification.
- Treat private-company funding, valuation and revenue references as funding/commercial signals, not as audited financial performance.
- Separate facts from interpretation. Do not make weak evidence sound definitive.
- Latest news/RSS headlines are external signals to verify, not proof on their own.
- Write for a smart business audience that may not be technical.
- Use clear, direct language and avoid AI hype.

Prioritise:
1. Mission, positioning and strategic focus
2. Executive leadership and why the leadership context matters
3. Products/services and what the company actually does
4. Funding/commercial signals, especially for private companies
5. Latest news only where it is relevant to the brief
6. Risks, caveats and questions for human review
7. Concrete next steps

Brief-type guidance:
- If the brief type is "Investment one-pager", emphasise market relevance, leadership, product surface, growth/funding signals, risks, and diligence next steps.
- If the brief type is "Company brief", emphasise what the company does, mission/positioning, leadership, key facts, recent signals, and business relevance.

SOURCE PACK:
{source_pack}

Return ONLY valid JSON. Do not include markdown, commentary, citations outside the JSON, or code fences.
Use exactly this schema and keep each field concise enough to fit a single PowerPoint slide:
{{
  "headline": "short title for the profile",
  "company_positioning": "1-2 sentences on mission, positioning and differentiation",
  "growth_direction": "1-2 sentences on likely growth direction or strategic direction, based only on the source pack",
  "target_market": "1-2 sentences on target users/customers/market",
  "company_description": "2-3 sentences suitable for a left-side profile panel",
  "what_they_do": ["specific bullet 1", "specific bullet 2", "specific bullet 3"],
  "leadership": ["Name — role; brief relevance", "Name — role; brief relevance", "Name — role; brief relevance"],
  "key_facts": ["specific fact 1", "specific fact 2", "specific fact 3", "specific fact 4"],
  "funding_commercial_signals": ["signal 1", "signal 2", "signal 3"],
  "latest_news_signals": ["recent signal 1", "recent signal 2", "recent signal 3"],
  "risks": ["risk / caveat 1", "risk / caveat 2", "risk / caveat 3"],
  "timeline": [
    {{"year": "2023", "text": "milestone or signal"}},
    {{"year": "2024", "text": "milestone or signal"}},
    {{"year": "2025", "text": "milestone or signal"}},
    {{"year": "2026", "text": "milestone, signal or to verify"}}
  ],
  "next_steps": ["specific next step 1", "specific next step 2", "specific next step 3"]
}}

Field-specific instructions:
- "leadership": include named executives only if they appear in the source pack. If leadership is missing, write "Leadership data to verify".
- "funding_commercial_signals": for private companies, use language such as "reported", "funding signal", "commercial signal", or "to verify" where appropriate.
- "latest_news_signals": include only relevant recent items. If the retrieved articles are weak, duplicated, irrelevant or missing, write that recent external signals require verification.
- "risks": include both information-quality risks and business risks where relevant.
- "next_steps": make these practical actions a human analyst or business team would take before using the one-pager.
""".strip()


def build_review_prompt(profile_json: str, source_pack: str) -> str:
    """Quality-control prompt for the second OpenAI call."""
    return f"""
You are reviewing an AI-generated company one-pager before it is used by a business team.

Compare the profile JSON against the source pack. Be strict and practical.

Return a concise review with four headings:

1. Claims to verify
- Flag any claim that appears unsupported, too broad, too confident, or dependent on a weak source.

2. Missing information
- Highlight any important missing details, especially leadership, mission, products, funding/commercial signals, financials, or recent news.

3. Risk notes
- Flag private-company financial uncertainty, duplicated/irrelevant RSS results, outdated data, or areas where human judgement is needed.

4. Suggested improvements for a non-technical audience
- Suggest how to make the one-pager clearer, more specific, or less jargon-heavy.

Rules:
- Do not rewrite the full one-pager.
- Do not introduce new facts.
- Be concise, direct and useful.
- Treat latest news as signals to verify, not definitive evidence.

SOURCE PACK:
{source_pack}

PROFILE JSON:
{profile_json}
""".strip()


# -----------------------------
# LLM provider layer
# -----------------------------
def get_secret(name: str) -> Optional[str]:
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




def run_llm(prompt: str, provider: str = "OpenAI") -> str:
    """Runs the generation or review prompt using OpenAI.

    If no OpenAI key is available, return an empty string so the app can fall back
    to the deterministic demo output instead of crashing during an interview demo.
    """
    if provider == "OpenAI" and get_secret("OPENAI_API_KEY"):
        return call_openai(prompt)
    return ""


# -----------------------------
# Profile parsing / formatting
# -----------------------------
def extract_json(text: str) -> Optional[Dict[str, Any]]:
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
        out = []
        for v in value:
            if isinstance(v, dict):
                text = v.get("text") or v.get("description") or " — ".join([str(x) for x in v.values() if x])
                out.append(str(text).strip())
            elif str(v).strip():
                out.append(str(v).strip())
        return out[:max_items]
    lines = [line.strip(" -•\t") for line in str(value).split("\n") if line.strip()]
    return lines[:max_items]


def df_to_bullets(df: pd.DataFrame, cols: List[str], max_items: int = 4) -> List[str]:
    if df.empty:
        return []
    bullets = []
    for _, row in df.head(max_items).iterrows():
        parts = [str(row.get(col, "")).strip() for col in cols if str(row.get(col, "")).strip()]
        bullets.append(" — ".join(parts))
    return bullets


def default_profile_sections(company: pd.Series, project: pd.Series, intel: Dict[str, pd.DataFrame], latest_articles: List[Dict[str, str]]) -> Dict[str, Any]:
    c = company.to_dict()
    financials = intel.get("financials", pd.DataFrame())
    f_row = financials.iloc[0].to_dict() if not financials.empty else {}
    milestones = []
    for _, row in intel.get("milestones", pd.DataFrame()).head(5).iterrows():
        milestones.append({"year": str(row.get("milestone_year", "")), "text": str(row.get("milestone_text", ""))})

    latest_news = []
    for article in latest_articles[:3]:
        title = article.get("title", "").strip()
        if title:
            latest_news.append(f"{title} — verify relevance before external use")
    if not latest_news:
        latest_news = ["Latest external news was excluded or unavailable.", "Use internal records and analyst notes as the primary evidence."]

    return {
        "headline": f"One Page Company Profile: {c.get('company_name', 'Company')}",
        "company_positioning": str(c.get("mission", "Mission not available.")),
        "growth_direction": str(c.get("expansion_signal", "Growth direction needs to be verified.")),
        "target_market": str(c.get("target_market", project.get("target_sector", "Target market to be confirmed."))),
        "company_description": str(c.get("short_description", f"{c.get('company_name', 'The company')} operates in {c.get('sector', 'its sector')}.")).strip(),
        "what_they_do": df_to_bullets(intel.get("products", pd.DataFrame()), ["product_name", "description"], 4) or [str(c.get("recent_news", "Products to verify."))],
        "leadership": df_to_bullets(intel.get("leadership", pd.DataFrame()), ["executive_name", "role"], 4) or ["Leadership not available in internal database."],
        "key_facts": [
            f"HQ: {c.get('hq_location', 'Not available')}",
            f"Founded: {c.get('founded_year', 'Not available')}",
            f"Type: {c.get('company_type', 'Not available')}",
            f"Relevance score: {c.get('relevance_score', 'Not available')}",
        ],
        "funding_commercial_signals": [
            str(f_row.get("latest_funding_or_valuation", "Funding/commercial data not available.")),
            str(f_row.get("commercial_signals", "Commercial signals require verification.")),
            str(f_row.get("financial_caveat", "Private-company financials should be treated cautiously.")),
        ],
        "latest_news_signals": latest_news,
        "risks": [
            "Private-company financials are incomplete unless the source explicitly discloses them.",
            "Latest RSS/news results may be duplicated, irrelevant or behind paywalls.",
            "Treat this as a first draft requiring human review before external use.",
        ],
        "timeline": milestones or [{"year": "TBC", "text": "Milestones not available in internal database."}],
        "next_steps": [
            "Verify leadership, funding and product details against current official sources.",
            "Check whether latest news is relevant and material.",
            "Refine the PowerPoint before sharing with stakeholders.",
        ],
    }


def normalise_profile_sections(raw_text: str, base: Dict[str, Any]) -> Dict[str, Any]:
    parsed = extract_json(raw_text) if raw_text else None
    if not parsed:
        return base
    for key in base:
        if key in parsed and parsed[key]:
            base[key] = parsed[key]
    return base


def profile_to_markdown(profile: Dict[str, Any]) -> str:
    def bullets(key: str, max_items: int = 5) -> str:
        return "\n".join([f"- {x}" for x in as_list(profile.get(key), max_items=max_items)])

    timeline_lines = []
    timeline = profile.get("timeline", [])
    if isinstance(timeline, list):
        for item in timeline[:5]:
            if isinstance(item, dict):
                timeline_lines.append(f"- **{item.get('year', '')}:** {item.get('text', '')}")
            else:
                timeline_lines.append(f"- {item}")

    return f"""
# {profile.get('headline', 'One Page Company Profile')}

## Mission & positioning
{profile.get('company_positioning', '')}

## Growth direction
{profile.get('growth_direction', '')}

## Target market
{profile.get('target_market', '')}

## Company description
{profile.get('company_description', '')}

## What they do
{bullets('what_they_do')}

## Executive leadership
{bullets('leadership')}

## Key facts
{bullets('key_facts')}

## Funding & commercial signals
{bullets('funding_commercial_signals')}

## Latest news / recent signals
{bullets('latest_news_signals')}

## Risks / things to verify
{bullets('risks')}

## Timeline / milestones
{chr(10).join(timeline_lines)}

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


def add_bullet_box(slide, x, y, w, h, items, font_size=8.2, color=None, max_items=4):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    items = as_list(items, max_items=max_items) or ["Not available / to verify."]
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = str(item)
        p.font.size = Pt(font_size)
        if color:
            p.font.color.rgb = color
        p.level = 0
        p.space_after = Pt(2)
    return box


def add_card(slide, x, y, w, h, title, body_items, icon_text="", max_items=4):
    navy = RGBColor(25, 52, 78)
    teal = RGBColor(79, 161, 190)
    dark = RGBColor(45, 55, 65)
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y + 0.05), Inches(0.42), Inches(0.42))
    circle.fill.solid(); circle.fill.fore_color.rgb = teal; circle.line.fill.background()
    add_textbox(slide, x + 0.07, y + 0.14, 0.28, 0.16, icon_text, font_size=7.5, bold=True, color=RGBColor(255, 255, 255), align=PP_ALIGN.CENTER)
    add_textbox(slide, x + 0.55, y, w - 0.55, 0.35, title, font_size=11.5, bold=True, color=navy)
    add_bullet_box(slide, x + 0.55, y + 0.42, w - 0.55, h - 0.42, body_items, font_size=7.8, color=dark, max_items=max_items)


def add_profile_pptx(profile: Dict[str, Any], company: pd.Series, project: pd.Series) -> BytesIO:
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

    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid(); bg.fill.fore_color.rgb = white; bg.line.fill.background()

    top_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.55))
    top_bg.fill.solid(); top_bg.fill.fore_color.rgb = light_blue; top_bg.line.fill.background()
    for x, width in [(5.7, 0.15), (6.35, 0.09), (7.1, 0.12), (7.8, 0.08), (8.45, 0.12), (9.4, 0.10), (10.0, 0.12), (10.9, 0.10), (11.7, 0.12)]:
        line = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(x), Inches(0), Inches(width), Inches(1.55))
        line.fill.solid(); line.fill.fore_color.rgb = RGBColor(175, 205, 220); line.line.fill.background()

    add_textbox(slide, 0.45, 0.25, 4.9, 0.45, "ONE PAGE COMPANY PROFILE", font_size=22, bold=False, color=navy)
    add_textbox(slide, 0.47, 0.74, 4.6, 0.2, str(profile.get("headline", "AI-generated business brief")), font_size=8.8, color=dark)

    ribbon = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(2.45), Inches(1.05), Inches(10.55), Inches(1.05))
    ribbon.fill.solid(); ribbon.fill.fore_color.rgb = deep_blue; ribbon.line.fill.background()

    top_sections = [
        (3.65, "M", "Mission / positioning", profile.get("company_positioning", "")),
        (6.75, "G", "Growth direction", profile.get("growth_direction", "")),
        (9.85, "T", "Target market", profile.get("target_market", "")),
    ]
    for x, icon, title, body in top_sections:
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(1.31), Inches(0.52), Inches(0.52))
        circ.fill.solid(); circ.fill.fore_color.rgb = teal; circ.line.fill.background()
        add_textbox(slide, x + 0.11, 1.43, 0.30, 0.18, icon, font_size=8.5, bold=True, color=white, align=PP_ALIGN.CENTER)
        add_textbox(slide, x + 0.65, 1.25, 2.1, 0.24, title, font_size=11.2, bold=True, color=white)
        add_textbox(slide, x + 0.65, 1.53, 2.25, 0.48, str(body), font_size=7.6, color=white)

    panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.48), Inches(2.35), Inches(2.85), Inches(4.65))
    panel.fill.solid(); panel.fill.fore_color.rgb = deep_blue; panel.line.fill.background()

    facts = [
        ("Company", company.get("company_name", "N/A")),
        ("HQ", company.get("hq_location", "N/A")),
        ("Founded", company.get("founded_year", "N/A")),
        ("Type", company.get("company_type", "N/A")),
        ("Sector", company.get("sector", "N/A")),
        ("Employees", company.get("employee_count", "N/A")),
        ("Score", company.get("relevance_score", "N/A")),
    ]
    y = 2.52
    for label, value in facts:
        add_textbox(slide, 0.66, y, 0.80, 0.22, label, font_size=7.2, bold=True, color=white)
        add_textbox(slide, 1.50, y, 1.55, 0.22, str(value), font_size=6.9, color=white, align=PP_ALIGN.RIGHT)
        divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.66), Inches(y + 0.28), Inches(2.38), Inches(0.01))
        divider.fill.solid(); divider.fill.fore_color.rgb = RGBColor(170, 190, 205); divider.line.fill.background()
        y += 0.38

    add_textbox(slide, 0.68, 5.55, 2.25, 0.3, "Company description", font_size=11.2, bold=True, color=white)
    add_textbox(slide, 0.68, 5.92, 2.25, 0.82, profile.get("company_description", ""), font_size=7.8, color=white)

    add_card(slide, 3.65, 2.35, 2.75, 1.20, "What they do", profile.get("what_they_do"), "W", max_items=3)
    add_card(slide, 6.85, 2.35, 2.75, 1.20, "Executive leadership", profile.get("leadership"), "E", max_items=4)
    add_card(slide, 10.05, 2.35, 2.80, 1.20, "Funding / commercial", profile.get("funding_commercial_signals"), "F", max_items=3)

    add_card(slide, 3.65, 3.88, 2.75, 1.25, "Latest news / signals", profile.get("latest_news_signals"), "L", max_items=3)
    add_card(slide, 6.85, 3.88, 2.75, 1.25, "Risks / verify", profile.get("risks"), "R", max_items=3)
    add_card(slide, 10.05, 3.88, 2.80, 1.25, "Next steps", profile.get("next_steps"), "N", max_items=3)

    add_textbox(slide, 3.65, 5.55, 3.3, 0.3, "Company milestones / signals", font_size=12.5, bold=True, color=navy)
    timeline_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.65), Inches(5.92), Inches(9.2), Inches(1.10))
    timeline_bg.fill.solid(); timeline_bg.fill.fore_color.rgb = light_grey; timeline_bg.line.color.rgb = mid_grey

    timeline = profile.get("timeline", [])
    if not isinstance(timeline, list) or not timeline:
        timeline = [{"year": "TBC", "text": "Milestones to verify"}]
    timeline = timeline[:5]
    x_positions = [4.15, 5.95, 7.75, 9.55, 11.35]
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.4), Inches(6.36), Inches(6.95), Inches(0.025))
    line.fill.solid(); line.fill.fore_color.rgb = teal; line.line.fill.background()
    for i, item in enumerate(timeline):
        if not isinstance(item, dict):
            item = {"year": str(2023 + i), "text": str(item)}
        x = x_positions[i]
        add_textbox(slide, x - 0.25, 6.05, 0.9, 0.25, item.get("year", str(2023 + i)), font_size=10.5, bold=True, color=navy, align=PP_ALIGN.CENTER)
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(6.30), Inches(0.16), Inches(0.16))
        dot.fill.solid(); dot.fill.fore_color.rgb = white; dot.line.color.rgb = teal; dot.line.width = Pt(2)
        add_textbox(slide, x - 0.48, 6.52, 1.12, 0.36, item.get("text", ""), font_size=6.5, color=dark, align=PP_ALIGN.CENTER)

    add_textbox(slide, 0.48, 7.18, 12.35, 0.16, "AI-generated first draft — funding, leadership and news signals require human verification before external use.", font_size=7.2, color=RGBColor(90, 90, 90), align=PP_ALIGN.RIGHT)

    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io


# -----------------------------
# Streamlit UI
# -----------------------------
st.title("AI Business Briefing Assistant")
st.caption("Prototype: free-text company request + internal company intelligence database + optional latest news + LLM workflow → editable PowerPoint one-pager")

if not DB_PATH.exists():
    try:
        import setup_demo_data
        setup_demo_data.main()
        DB_PATH = resolve_db_path()
        st.info("Demo database was missing, so it has been created automatically.")
    except Exception as e:
        st.error(f"Demo database could not be created automatically: {e}")
        st.stop()

companies = safe_load_table("companies", str(DB_PATH))
if companies.empty:
    st.error("No companies table found in the demo database. Re-upload briefing_demo.db or setup_demo_data.py.")
    st.stop()
notes = read_markdown_notes()

with st.sidebar:
    st.header("Demo controls")
    provider = st.selectbox("LLM mode", ["OpenAI", "Demo fallback"])
    if provider == "OpenAI" and not get_secret("OPENAI_API_KEY"):
        st.warning("OpenAI key not found. Add OPENAI_API_KEY in Streamlit secrets, or use Demo fallback.")
    st.divider()
    st.caption(f"Database path: {DB_PATH.relative_to(BASE) if DB_PATH.exists() else DB_PATH}")
    st.caption(f"Internal company records: {len(companies)}")
    st.caption(f"Markdown notes found: {len(notes)}")
    for table in ["leadership", "products", "financials", "milestones"]:
        st.caption(f"{table.title()} records: {len(safe_load_table(table, str(DB_PATH)))}")

st.subheader("1. Enter company and brief type")
company_query = st.text_input("Company name", value="Anthropic", placeholder="Type any company name, for example Anthropic, OpenAI or Mistral AI")
brief_type = st.selectbox("Brief type", ["Investment one-pager", "Company brief"])

st.subheader("2. Add context")
extra_notes = st.text_area(
    "Additional context, messy notes or instructions",
    height=150,
    placeholder="Example: Focus on enterprise AI adoption, leadership, funding/commercial signals, risks, and why this company matters for a senior stakeholder.",
)
uploaded_file = st.file_uploader("Optional: upload a .txt or .md file", type=["txt", "md"])
if uploaded_file is not None:
    extra_notes += "\n\nUPLOADED FILE CONTENT:\n" + uploaded_file.read().decode("utf-8", errors="ignore")

st.subheader("3. Match to internal database")
selected_company, has_internal_match, match_message = find_company_match(company_query, companies)
if has_internal_match:
    st.success(match_message)
else:
    st.warning(match_message)

intel = company_intelligence(selected_company) if has_internal_match else {"leadership": pd.DataFrame(), "products": pd.DataFrame(), "financials": pd.DataFrame(), "milestones": pd.DataFrame()}
with st.expander("View matched internal company intelligence"):
    st.write("Company profile")
    st.dataframe(pd.DataFrame(selected_company).rename(columns={0: "value"}))
    for label, df in intel.items():
        st.write(label.title())
        if df.empty:
            st.caption("No internal records found.")
        else:
            st.dataframe(df)

st.subheader("4. Latest news signals")
include_latest_news = st.checkbox(
    "Include latest 5 news/RSS articles in the one-pager",
    value=True,
    help="If selected, the app retrieves recent Google News RSS results for the company and passes them into the LLM as external signals to verify.",
)
if include_latest_news:
    st.info("Latest news will be retrieved when you generate the one-pager. It is a signal source, not definitive evidence.")
else:
    st.info("Latest news is excluded. The one-pager will rely on internal data, notes and your added context.")

st.subheader("5. Generate output")
st.write("This runs the workflow: interpret request → match internal company intelligence → retrieve notes → optional news lookup → draft structured profile → review → export editable PowerPoint.")

if st.button("Generate one-pager", type="primary"):
    selected_project = build_project_context(brief_type, extra_notes)
    query = f"{company_query} {selected_company.get('company_name', '')} {brief_type} {extra_notes}"
    retrieved_notes = retrieve_relevant_notes(query, notes)

    with st.spinner("Retrieving latest news signals..." if include_latest_news else "Preparing source pack..."):
        latest_articles = get_latest_company_articles(selected_company.get("company_name", company_query), max_articles=5) if include_latest_news else []
        source_pack = build_source_pack(selected_company, selected_project, extra_notes, retrieved_notes, latest_articles, intel)

    base_profile = default_profile_sections(selected_company, selected_project, intel, latest_articles)
    with st.spinner("Drafting structured profile..."):
        if provider == "Demo fallback":
            raw_draft = json.dumps(base_profile, indent=2)
            profile = base_profile
        else:
            draft_prompt = build_prompt(source_pack, brief_type)
            raw_draft = run_llm(draft_prompt, provider)
            profile = normalise_profile_sections(raw_draft, base_profile)
        markdown_draft = profile_to_markdown(profile)

    with st.spinner("Running review step..."):
        if provider == "Demo fallback":
            review = "Demo review: verify all leadership, funding/valuation and latest-news claims before external use. Treat private-company financial data as directional commercial signals, not full financial performance."
        else:
            review_prompt = build_review_prompt(json.dumps(profile, indent=2), source_pack)
            review = run_llm(review_prompt, provider)

    pptx_file = add_profile_pptx(profile, selected_company, selected_project)
    clean_name = re.sub(r"[^A-Za-z0-9_]+", "_", str(selected_company.get("company_name", company_query))).strip("_") or "company"
    st.session_state["raw_draft"] = raw_draft
    st.session_state["profile"] = profile
    st.session_state["draft"] = markdown_draft
    st.session_state["review"] = review
    st.session_state["source_pack"] = source_pack
    st.session_state["latest_articles"] = latest_articles
    st.session_state["include_latest_news"] = include_latest_news
    st.session_state["match_message"] = match_message
    st.session_state["pptx_file"] = pptx_file.getvalue()
    st.session_state["pptx_name"] = f"{clean_name}_{brief_type.replace(' ', '_').lower()}.pptx"

if "profile" in st.session_state:
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["PowerPoint output", "Text preview", "Latest news", "AI review", "Source pack"])
    with tab1:
        st.success("PowerPoint one-pager created. Download and open in PowerPoint to edit the slide.")
        st.caption(st.session_state.get("match_message", ""))
        st.download_button(
            "Download one-page profile as PowerPoint",
            data=st.session_state["pptx_file"],
            file_name=st.session_state.get("pptx_name", "one_page_profile.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        with st.expander("View structured fields used in the slide"):
            st.json(st.session_state["profile"])
    with tab2:
        st.markdown(st.session_state["draft"])
        st.download_button("Download text preview as Markdown", data=st.session_state["draft"], file_name="ai_one_pager_preview.md", mime="text/markdown")
    with tab3:
        if not st.session_state.get("include_latest_news"):
            st.write("Latest news was excluded for this run.")
        elif st.session_state.get("latest_articles"):
            for article in st.session_state.get("latest_articles", []):
                st.markdown(f"**{article.get('title', '')}**")
                if article.get("published"):
                    st.caption(article.get("published"))
                if article.get("summary"):
                    st.write(article.get("summary"))
                if article.get("link"):
                    st.write(article.get("link"))
                st.divider()
        else:
            st.write("No latest articles retrieved for this company.")
    with tab4:
        st.markdown(st.session_state["review"])
    with tab5:
        st.code(st.session_state["source_pack"], language="text")

st.divider()
st.caption(
    "Interview talking point: this version starts with a free-text company request, checks a richer internal company intelligence database "
    "for profile, mission, leadership, products, funding/commercial signals and milestones, optionally retrieves latest external signals, "
    "then uses an LLM workflow to structure the evidence into an editable PowerPoint one-pager."
)
