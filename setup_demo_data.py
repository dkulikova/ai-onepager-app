import shutil
import sqlite3
from pathlib import Path

BASE = Path(__file__).parent
DATA_DIR = BASE / "data"
NOTES_DIR = BASE / "notes"
DB_PATH = DATA_DIR / "briefing_demo.db"

companies = [
    {
        "company_id": 1,
        "company_name": "OpenAI",
        "sector": "Frontier AI / AI products and platform",
        "hq_location": "San Francisco, US",
        "founded_year": 2015,
        "company_type": "Private; capped-profit structure controlled by nonprofit parent",
        "employee_count": "Not publicly confirmed; fast-growing global workforce",
        "website": "https://openai.com/",
        "current_office_location": "San Francisco HQ with global offices and hiring across major technology and policy markets",
        "mission": "Ensure that artificial general intelligence benefits all of humanity.",
        "short_description": "OpenAI is an AI research and deployment company best known for ChatGPT, the OpenAI API, frontier models, enterprise AI products and developer tooling.",
        "recent_news": "OpenAI continues to expand its business platform across ChatGPT, API access, coding tools, agents and enterprise deployment.",
        "expansion_signal": "Strong demand across consumer, developer and enterprise channels creates ongoing need for product, infrastructure, sales, policy, safety and customer-facing teams.",
        "target_market": "Consumers, developers, enterprises, startups, governments and builders deploying AI into products and workflows.",
        "differentiation": "Broadest recognition in consumer AI, deep developer ecosystem, model/API platform and enterprise product surface.",
        "relevance_score": 95,
    },
    {
        "company_id": 2,
        "company_name": "Anthropic",
        "sector": "AI safety / Claude enterprise AI",
        "hq_location": "San Francisco, US",
        "founded_year": 2021,
        "company_type": "Private Public Benefit Corporation",
        "employee_count": "Not publicly confirmed; rapidly scaling research, product and go-to-market organisation",
        "website": "https://www.anthropic.com/",
        "current_office_location": "San Francisco HQ with expanding US and European presence including London, Dublin, Zurich, Paris and Munich",
        "mission": "Build reliable, interpretable and steerable AI systems that people can rely on.",
        "short_description": "Anthropic is an AI safety and research company building Claude, Claude Enterprise, Claude API and related enterprise and developer products.",
        "recent_news": "Anthropic continues to expand Claude across enterprise, developer, coding, security and productivity use cases.",
        "expansion_signal": "Enterprise demand for Claude, international office expansion and product growth indicate a scaling commercial and research organisation.",
        "target_market": "Enterprises, developers, regulated organisations, researchers and teams seeking reliable AI assistants and AI agents.",
        "differentiation": "Safety-led positioning, enterprise reliability, model governance and emphasis on interpretable and steerable AI systems.",
        "relevance_score": 93,
    },
    {
        "company_id": 3,
        "company_name": "Mistral AI",
        "sector": "Frontier AI / open models and enterprise AI platform",
        "hq_location": "Paris, France",
        "founded_year": 2023,
        "company_type": "Private",
        "employee_count": "Not publicly confirmed; fast-growing European AI scale-up",
        "website": "https://mistral.ai/",
        "current_office_location": "Paris HQ with international presence including London, Palo Alto, Germany and Singapore",
        "mission": "Democratise frontier AI through open, portable and enterprise-ready AI systems.",
        "short_description": "Mistral AI is a Paris-headquartered frontier AI company focused on open models, enterprise AI, custom assistants, agents and deployment flexibility.",
        "recent_news": "Mistral is scaling its enterprise AI platform, Le Chat/Vibe product family and infrastructure partnerships.",
        "expansion_signal": "Open-model strategy, European sovereign AI positioning and enterprise platform focus create strong international growth signals.",
        "target_market": "Enterprises, developers, public-sector organisations and companies needing custom, portable or self-deployable AI systems.",
        "differentiation": "European AI champion with open-model strategy, deployment portability and enterprise-grade customisation.",
        "relevance_score": 88,
    },
]

leadership = [
    (1, "Sam Altman", "CEO and co-founder", "Leads OpenAI's overall company strategy and AGI mission."),
    (1, "Greg Brockman", "President and co-founder", "Senior technical and organisational leader involved from OpenAI's founding."),
    (1, "Fidji Simo", "CEO of Applications", "Leads OpenAI's applications organisation, product and business execution."),
    (1, "Sarah Friar", "Chief Financial Officer", "Finance leader supporting OpenAI's scaling and commercial operations."),
    (1, "Jakub Pachocki", "Chief Scientist", "Research leader responsible for frontier model development."),
    (2, "Dario Amodei", "CEO and co-founder", "Leads Anthropic's research-led safety and product strategy."),
    (2, "Daniela Amodei", "President and co-founder", "Senior leader and board member focused on company operations and mission execution."),
    (2, "Mike Krieger", "Chief Product Officer", "Product leader driving Claude's product and enterprise experience."),
    (2, "Jared Kaplan", "Co-founder / research leader", "Co-founder associated with Anthropic's technical and research direction."),
    (3, "Arthur Mensch", "Co-founder and CEO", "Leads Mistral AI and its European frontier AI strategy."),
    (3, "Guillaume Lample", "Co-founder and Chief Science Officer", "Research leader focused on model development and scientific direction."),
    (3, "Timothée Lacroix", "Co-founder and CTO", "Technology leader focused on engineering and platform development."),
]

products = [
    (1, "ChatGPT", "AI assistant", "Consumer and business AI assistant used for writing, analysis, coding and productivity."),
    (1, "ChatGPT Enterprise / Business", "Enterprise AI", "Workplace AI offering for organisations with admin, privacy and collaboration controls."),
    (1, "OpenAI API", "Developer platform", "Model access for developers and enterprises building AI-powered products and workflows."),
    (1, "Codex", "Coding assistant", "AI coding product for software development and engineering workflows."),
    (1, "Sora", "Media generation", "Video generation product family and creative AI capability."),
    (2, "Claude", "AI assistant", "Anthropic's assistant for writing, analysis, coding and business workflows."),
    (2, "Claude Enterprise", "Enterprise AI", "Organization-wide Claude deployment with governance, security and admin capabilities."),
    (2, "Claude API", "Developer platform", "Programmatic access to Claude models for developers and enterprise applications."),
    (2, "Claude Code", "Coding assistant", "Agentic coding product for development teams."),
    (2, "Claude for Microsoft 365", "Workplace integration", "Claude integrated into business productivity workflows."),
    (3, "Vibe / Le Chat", "AI assistant and agent", "Mistral's chat and work/coding assistant product."),
    (3, "La Plateforme / Studio", "Developer and enterprise platform", "Platform for building, deploying and managing AI apps and agents."),
    (3, "Forge", "Custom model development", "Tools and services for custom model development and fine-tuning."),
    (3, "Open models", "Model portfolio", "Mistral model family for open, portable and custom deployments."),
    (3, "Le Chat Enterprise", "Enterprise AI", "Private and customisable AI productivity platform for teams."),
]

financials = [
    {
        "company_id": 1,
        "ownership_status": "Private; no public audited financial statements available",
        "latest_funding_or_valuation": "OpenAI announced a 2026 funding round with $122bn in committed capital at a post-money valuation of $852bn.",
        "commercial_signals": "Large-scale consumer usage, enterprise adoption, API ecosystem, coding tools and infrastructure partnerships.",
        "investors_partners": "Microsoft, SoftBank, Nvidia and other strategic/financial investors have been publicly associated with OpenAI funding or partnerships.",
        "financial_caveat": "Revenue, profitability and valuation figures are time-sensitive; use official announcements or reputable financial sources before external use.",
    },
    {
        "company_id": 2,
        "ownership_status": "Private Public Benefit Corporation; no public audited financial statements available",
        "latest_funding_or_valuation": "Anthropic announced a 2026 Series H funding round of $65bn at a $965bn post-money valuation.",
        "commercial_signals": "Rapid Claude enterprise/API adoption, product expansion into coding, security, productivity and sector-specific workflows.",
        "investors_partners": "Publicly reported investors and partners include major cloud, strategic and institutional backers; verify current cap table before external use.",
        "financial_caveat": "Financial performance is not publicly reported like a listed company; funding and valuation are signals, not complete financial performance.",
    },
    {
        "company_id": 3,
        "ownership_status": "Private; no public audited financial statements available",
        "latest_funding_or_valuation": "Reuters reported Mistral AI raised €1.7bn in Series C funding in 2025 at an €11.7bn valuation, led by ASML.",
        "commercial_signals": "Enterprise platform growth, strategic European AI positioning, custom deployments, open-model adoption and AI infrastructure investment.",
        "investors_partners": "ASML, Nvidia, Andreessen Horowitz, Bpifrance, General Catalyst, Index Ventures and other investors have been publicly reported.",
        "financial_caveat": "Revenue and customer figures are not consistently disclosed; treat third-party estimates as directional only.",
    },
]

milestones = [
    (1, "2015", "OpenAI founded."),
    (1, "2022", "ChatGPT launched publicly."),
    (1, "2024", "Enterprise and multimodal AI adoption accelerated."),
    (1, "2025", "Applications leadership expanded with Fidji Simo appointment."),
    (1, "2026", "Major funding and enterprise/platform expansion signal continued scale-up."),
    (2, "2021", "Anthropic founded by former OpenAI researchers and leaders."),
    (2, "2023", "Claude product family gained wider enterprise and developer attention."),
    (2, "2024", "Claude Enterprise and API adoption expanded."),
    (2, "2025", "European footprint and Claude product portfolio expanded."),
    (2, "2026", "Large funding round and continued enterprise product expansion."),
    (3, "2023", "Mistral AI founded in Paris by former DeepMind and Meta researchers."),
    (3, "2024", "Open model releases and enterprise platform visibility accelerated."),
    (3, "2025", "Le Chat Enterprise launched and ASML-led Series C reported."),
    (3, "2026", "Reuters reported new debt financing for AI data-centre build-out."),
]

projects = [
    {
        "project_name": "Investment one-pager",
        "client": "Interview demo",
        "location": "Global / enterprise AI market",
        "target_sector": "Frontier AI labs, enterprise AI platforms, AI assistants and AI infrastructure",
        "brief_description": "Create a concise investment-style profile covering mission, leadership, products, funding/commercial signals, market relevance, risks and next diligence steps.",
    },
    {
        "project_name": "Company brief",
        "client": "Senior stakeholder meeting",
        "location": "UK, Europe and US",
        "target_sector": "Enterprise AI adoption, AI agents, coding assistants and AI productivity platforms",
        "brief_description": "Create a business-ready one-pager on an AI company using structured company intelligence, analyst notes and optional latest news signals.",
    },
]

notes = {
    "openai.md": """# OpenAI research notes\n\nMission & positioning:\nOpenAI describes itself as an AI research and deployment company. Its mission is to ensure that artificial general intelligence benefits all of humanity.\n\nProducts and commercial context:\n- ChatGPT is the most recognisable user-facing product.\n- ChatGPT Enterprise / Business and the OpenAI API are core enterprise and developer channels.\n- Codex expands OpenAI's relevance in software engineering and agentic coding workflows.\n- Sora and multimodal products extend the company into creative and media generation.\n\nLeadership notes:\n- Sam Altman is CEO and co-founder.\n- Greg Brockman is president and co-founder.\n- Fidji Simo joined as CEO of Applications, reporting to Sam Altman.\n- Sarah Friar is CFO and Jakub Pachocki is Chief Scientist.\n\nCaveats:\n- OpenAI is private, so full financial performance is not reported like a public company.\n- Revenue, valuation and product details change quickly and should be checked before external use.\n\nManual review links:\n- https://openai.com/about/\n- https://chatgpt.com/business/enterprise/\n- https://openai.com/api/\n""",
    "anthropic.md": """# Anthropic research notes\n\nMission & positioning:\nAnthropic is an AI safety and research company building reliable, interpretable and steerable AI systems. Its main product family is Claude.\n\nProducts and commercial context:\n- Claude is used for writing, coding, analysis and enterprise workflows.\n- Claude Enterprise adds governance, admin and security controls for organisations.\n- Claude API supports developers building AI applications.\n- Claude Code and sector-specific Claude products expand into more specialised workflows.\n\nLeadership notes:\n- Dario Amodei is CEO and co-founder.\n- Daniela Amodei is President and co-founder.\n- Mike Krieger is Chief Product Officer.\n\nCaveats:\n- Anthropic is private, so financial performance is not fully public.\n- Safety and enterprise trust are central differentiators but should be tied to concrete evidence, not treated as marketing claims alone.\n\nManual review links:\n- https://www.anthropic.com/company\n- https://www.anthropic.com/product/enterprise\n- https://platform.claude.com/docs/\n""",
    "mistral_ai.md": """# Mistral AI research notes\n\nMission & positioning:\nMistral AI is a Paris-headquartered frontier AI company focused on open models, enterprise AI, custom assistants and deployment flexibility. It is a useful European counterpoint to US-based AI labs.\n\nProducts and commercial context:\n- Vibe / Le Chat is Mistral's AI assistant and agent product for work and code.\n- La Plateforme / Studio supports custom agents, AI apps, observability and deployment control.\n- Forge supports custom model development and fine-tuning.\n- Mistral emphasises deployment portability from edge to cloud.\n\nLeadership notes:\n- Arthur Mensch is co-founder and CEO.\n- Guillaume Lample is co-founder and Chief Science Officer.\n- Timothée Lacroix is co-founder and CTO.\n\nCaveats:\n- Product naming has changed over time, so older references to Le Chat may map to Vibe.\n- Revenue and customer figures are not fully disclosed and should be treated carefully.\n\nManual review links:\n- https://mistral.ai/\n- https://mistral.ai/about/\n- https://mistral.ai/news/le-chat-enterprise/\n""",
}

old_note_names = ["northstar_robotics.md", "helio_grid.md", "openai.md", "anthropic.md", "mistral_ai.md"]


def main() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    NOTES_DIR.mkdir(exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    for table in ["companies", "projects", "leadership", "products", "financials", "milestones"]:
        cur.execute(f"DROP TABLE IF EXISTS {table}")

    cur.execute("""
        CREATE TABLE companies (
            company_id INTEGER PRIMARY KEY,
            company_name TEXT,
            sector TEXT,
            hq_location TEXT,
            founded_year INTEGER,
            company_type TEXT,
            employee_count TEXT,
            website TEXT,
            current_office_location TEXT,
            mission TEXT,
            short_description TEXT,
            recent_news TEXT,
            expansion_signal TEXT,
            target_market TEXT,
            differentiation TEXT,
            relevance_score INTEGER
        )
    """)
    cur.execute("""
        CREATE TABLE leadership (
            leader_id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            executive_name TEXT,
            role TEXT,
            short_bio TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE products (
            product_id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            product_name TEXT,
            product_category TEXT,
            description TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE financials (
            financial_id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            ownership_status TEXT,
            latest_funding_or_valuation TEXT,
            commercial_signals TEXT,
            investors_partners TEXT,
            financial_caveat TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE milestones (
            milestone_id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            milestone_year TEXT,
            milestone_text TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE projects (
            project_id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_name TEXT,
            client TEXT,
            location TEXT,
            target_sector TEXT,
            brief_description TEXT
        )
    """)

    for row in companies:
        cur.execute("""
            INSERT INTO companies
            (company_id, company_name, sector, hq_location, founded_year, company_type, employee_count, website, current_office_location,
             mission, short_description, recent_news, expansion_signal, target_market, differentiation, relevance_score)
            VALUES (:company_id, :company_name, :sector, :hq_location, :founded_year, :company_type, :employee_count, :website, :current_office_location,
                    :mission, :short_description, :recent_news, :expansion_signal, :target_market, :differentiation, :relevance_score)
        """, row)

    cur.executemany("INSERT INTO leadership (company_id, executive_name, role, short_bio) VALUES (?, ?, ?, ?)", leadership)
    cur.executemany("INSERT INTO products (company_id, product_name, product_category, description) VALUES (?, ?, ?, ?)", products)
    for row in financials:
        cur.execute("""
            INSERT INTO financials (company_id, ownership_status, latest_funding_or_valuation, commercial_signals, investors_partners, financial_caveat)
            VALUES (:company_id, :ownership_status, :latest_funding_or_valuation, :commercial_signals, :investors_partners, :financial_caveat)
        """, row)
    cur.executemany("INSERT INTO milestones (company_id, milestone_year, milestone_text) VALUES (?, ?, ?)", milestones)
    for row in projects:
        cur.execute("""
            INSERT INTO projects (project_name, client, location, target_sector, brief_description)
            VALUES (:project_name, :client, :location, :target_sector, :brief_description)
        """, row)

    conn.commit()
    conn.close()

    for name in old_note_names:
        for folder in [NOTES_DIR, BASE]:
            path = folder / name
            if path.exists():
                path.unlink()

    for name, content in notes.items():
        (NOTES_DIR / name).write_text(content, encoding="utf-8")
        (BASE / name).write_text(content, encoding="utf-8")

    shutil.copy2(DB_PATH, BASE / "briefing_demo.db")

    print(f"Demo database created at: {DB_PATH}")
    print(f"Flat upload copy created at: {BASE / 'briefing_demo.db'}")
    print(f"Demo notes created at: {NOTES_DIR}")


if __name__ == "__main__":
    main()
