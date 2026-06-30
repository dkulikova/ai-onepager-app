import shutil
import sqlite3
from pathlib import Path

BASE = Path(__file__).parent
DATA_DIR = BASE / "data"
NOTES_DIR = BASE / "notes"
DB_PATH = DATA_DIR / "briefing_demo.db"

# Demo records use real companies, but values are intentionally concise and presentation-ready.
# Employee counts are approximate public/indicative figures for prototype purposes.
companies = [
    {
        "company_name": "OpenAI",
        "sector": "Frontier AI / AI products and platform",
        "hq_location": "San Francisco, US",
        "employee_count": 5000,
        "current_office_location": "San Francisco HQ with international hiring/office footprint including London, Dublin, Washington DC, Singapore and Tokyo",
        "recent_news": "OpenAI positions its business offering as an enterprise AI platform spanning ChatGPT, Codex and API access, with tools for work, coding, agents and internal knowledge.",
        "expansion_signal": "Strong enterprise adoption and broad product surface across ChatGPT, API and Codex create demand for sales, customer success, engineering and workplace operations across major markets.",
        "relevance_score": 95,
    },
    {
        "company_name": "Anthropic",
        "sector": "AI safety / Claude enterprise AI",
        "hq_location": "San Francisco, US",
        "employee_count": 2500,
        "current_office_location": "San Francisco HQ with expanding offices across the US and Europe, including London, Dublin, Zurich, Paris and Munich",
        "recent_news": "Anthropic positions Claude Enterprise around organization-wide deployment, governance, data controls and admin infrastructure for enterprise IT and security teams.",
        "expansion_signal": "Rapid enterprise demand for Claude, international expansion and growing European footprint indicate a scaling commercial and research organisation.",
        "relevance_score": 93,
    },
    {
        "company_name": "Mistral AI",
        "sector": "Frontier AI / open models and enterprise AI platform",
        "hq_location": "Paris, France",
        "employee_count": 350,
        "current_office_location": "Paris HQ with international presence including London, Palo Alto, Germany and Singapore",
        "recent_news": "Mistral AI markets itself as an enterprise AI platform for custom assistants, autonomous agents, multimodal AI and deployment portability from edge to cloud.",
        "expansion_signal": "Its open-model strategy, enterprise platform focus and international office footprint make it a strong example of a European AI scale-up expanding globally.",
        "relevance_score": 88,
    },
]

projects = [
    {
        "project_name": "AI Company Profile",
        "client": "Interview demo",
        "location": "Global / enterprise AI market",
        "target_sector": "Frontier AI labs and enterprise AI platforms",
        "brief_description": "Create a concise one-page company profile that explains positioning, growth signals, market relevance and caveats for a non-technical business audience.",
    },
    {
        "project_name": "Executive AI Briefing",
        "client": "Senior stakeholder meeting",
        "location": "UK, Europe and US",
        "target_sector": "Enterprise AI adoption, AI agents, coding assistants and AI productivity platforms",
        "brief_description": "Prepare a business-ready one-pager on an AI company, using structured records and analyst notes to generate a polished PowerPoint output.",
    },
]

openai_note = """# OpenAI research notes

Positioning:
OpenAI is an AI research and deployment company. Its stated mission is to ensure that artificial general intelligence benefits all of humanity. For business users, the company is best known for ChatGPT, ChatGPT Enterprise/Business, Codex and the OpenAI API platform.

Useful evidence for a one-pager:
- OpenAI's business offering is positioned as a complete AI platform spanning ChatGPT, Codex and API access.
- Enterprise messaging focuses on productivity, internal knowledge, agents, coding and team-wide deployment.
- OpenAI careers listings show hiring across San Francisco and international markets, including London, Dublin, Washington DC, Singapore and Tokyo.
- The company is a useful demo profile because it is highly recognisable to non-technical audiences.

Caveats / things to verify:
- Employee count and office footprint change quickly and should be treated as indicative.
- Product names, models and enterprise packaging can change frequently.
- For external use, validate any commercial claims directly against OpenAI's current official pages.

Source links for manual review:
- https://openai.com/about/
- https://openai.com/business/
- https://openai.com/api/
- https://openai.com/careers/search/
"""

anthropic_note = """# Anthropic research notes

Positioning:
Anthropic is an AI safety and research company focused on building reliable, interpretable and steerable AI systems. Its main product family is Claude, including Claude for individuals, teams, enterprises and API developers.

Useful evidence for a one-pager:
- Claude Enterprise is positioned around organization-wide deployment with governance, data controls and admin infrastructure for IT and security teams.
- Anthropic's brand is differentiated by safety, reliability and enterprise trust.
- Public announcements describe expansion across Europe, including offices in London, Dublin, Zurich, Paris and Munich.
- Anthropic is a strong demo profile because it lets the presenter explain AI safety, enterprise deployment and human oversight in plain language.

Caveats / things to verify:
- Office expansion and employee counts are highly time-sensitive.
- The Claude model family and product tiers change quickly.
- Use current Anthropic sources before quoting specific model capabilities or customer figures.

Source links for manual review:
- https://www.anthropic.com/
- https://www.anthropic.com/product/enterprise
- https://www.anthropic.com/news/new-offices-in-paris-and-munich-expand-european-presence
- https://platform.claude.com/docs/en/api/overview
"""

mistral_note = """# Mistral AI research notes

Positioning:
Mistral AI is a Paris-headquartered frontier AI company focused on open models, enterprise AI, custom agents and deployment flexibility. It is useful as a European counterpoint to US-based AI labs.

Useful evidence for a one-pager:
- Mistral describes its platform as enabling enterprises to customize, fine-tune and deploy AI assistants, autonomous agents and multimodal AI.
- The company emphasizes deployment portability, including running AI from edge to cloud.
- Vibe, formerly Le Chat, is positioned as Mistral's AI chat and agent product for work and code.
- Public company profiles describe a growing international footprint across Paris, London, Palo Alto, Germany and Singapore.

Caveats / things to verify:
- Employee count is approximate and changes as the company scales.
- Product naming has changed, so older references to Le Chat may now map to Vibe.
- As with any fast-growing AI company, commercial positioning and product packaging evolve quickly.

Source links for manual review:
- https://mistral.ai/
- https://mistral.ai/products/vibe/
- https://mistral.ai/news/le-chat-enterprise
- https://www.welcometothejungle.com/en/companies/mistral-ai/team-1
"""

notes = {
    "openai.md": openai_note,
    "anthropic.md": anthropic_note,
    "mistral_ai.md": mistral_note,
}

old_note_names = [
    "northstar_robotics.md",
    "helio_grid.md",
    "openai.md",
    "anthropic.md",
    "mistral_ai.md",
]


def main() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    NOTES_DIR.mkdir(exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("DROP TABLE IF EXISTS companies")
    cur.execute("DROP TABLE IF EXISTS projects")
    cur.execute(
        """
        CREATE TABLE companies (
            company_id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT,
            sector TEXT,
            hq_location TEXT,
            employee_count INTEGER,
            current_office_location TEXT,
            recent_news TEXT,
            expansion_signal TEXT,
            relevance_score INTEGER
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE projects (
            project_id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_name TEXT,
            client TEXT,
            location TEXT,
            target_sector TEXT,
            brief_description TEXT
        )
        """
    )
    for row in companies:
        cur.execute(
            """
            INSERT INTO companies
            (company_name, sector, hq_location, employee_count, current_office_location, recent_news, expansion_signal, relevance_score)
            VALUES (:company_name, :sector, :hq_location, :employee_count, :current_office_location, :recent_news, :expansion_signal, :relevance_score)
            """,
            row,
        )
    for row in projects:
        cur.execute(
            """
            INSERT INTO projects
            (project_name, client, location, target_sector, brief_description)
            VALUES (:project_name, :client, :location, :target_sector, :brief_description)
            """,
            row,
        )
    conn.commit()
    conn.close()

    # Remove old demo notes from prior version, then write current notes.
    for name in old_note_names:
        for folder in [NOTES_DIR, BASE]:
            path = folder / name
            if path.exists():
                path.unlink()

    for name, content in notes.items():
        (NOTES_DIR / name).write_text(content, encoding="utf-8")
        # Convenience copies for manual GitHub web uploads where everything is uploaded flat.
        (BASE / name).write_text(content, encoding="utf-8")

    # Convenience copy for manual GitHub web uploads where everything is uploaded flat.
    shutil.copy2(DB_PATH, BASE / "briefing_demo.db")

    print(f"Demo database created at: {DB_PATH}")
    print(f"Flat upload copy created at: {BASE / 'briefing_demo.db'}")
    print(f"Demo notes created at: {NOTES_DIR}")


if __name__ == "__main__":
    main()
