import shutil
import sqlite3
from pathlib import Path

BASE = Path(__file__).parent
DATA_DIR = BASE / "data"
NOTES_DIR = BASE / "notes"
DB_PATH = DATA_DIR / "briefing_demo.db"

companies = [
    {
        "company_name": "Northstar Robotics",
        "sector": "AI / Robotics",
        "hq_location": "London, UK",
        "employee_count": 420,
        "current_office_location": "King's Cross, London",
        "recent_news": "Announced a new enterprise robotics platform and opened a European customer success hub.",
        "expansion_signal": "Hiring across sales, implementation and product roles; recent move into larger enterprise contracts.",
        "relevance_score": 86,
    },
    {
        "company_name": "HelioGrid Energy",
        "sector": "Climate Tech / Energy Software",
        "hq_location": "Manchester, UK",
        "employee_count": 260,
        "current_office_location": "Manchester with small London client office",
        "recent_news": "Raised Series B funding to expand grid analytics and utility partnerships.",
        "expansion_signal": "Likely need for a larger London presence due to client concentration and investor meetings.",
        "relevance_score": 78,
    },
    {
        "company_name": "Aster Health AI",
        "sector": "Health AI",
        "hq_location": "Cambridge, UK",
        "employee_count": 180,
        "current_office_location": "Cambridge Science Park",
        "recent_news": "Signed pilots with two NHS trusts and started recruiting commercial leadership.",
        "expansion_signal": "Commercial scaling phase; potential need for London office to support partnerships.",
        "relevance_score": 74,
    },
]

projects = [
    {
        "project_name": "AI Occupier Brief",
        "client": "Internal leasing team",
        "location": "Central London",
        "target_sector": "AI, software, robotics, climate tech",
        "brief_description": "Create a concise briefing note for a leasing opportunity focused on fast-growing technology occupiers.",
    },
    {
        "project_name": "Executive Client Prep",
        "client": "Senior leadership meeting",
        "location": "UK-wide",
        "target_sector": "High-growth companies",
        "brief_description": "Prepare a one-page summary of relevant companies, key signals and recommended discussion points.",
    },
]

northstar_note = """# Northstar Robotics research notes

Northstar Robotics is moving from product-led growth into enterprise deployments. Public hiring signals suggest a stronger focus on sales engineering, implementation, customer success and partnerships.

Useful evidence for a one-pager:
- The company is likely to need client-facing space if enterprise customers become a larger part of revenue.
- The current office location is convenient for talent but may not be ideal for executive meetings.
- Risk: expansion interest is inferred from hiring and funding signals; there is no confirmed property search.
"""

helio_note = """# HelioGrid Energy research notes

HelioGrid sells analytics software to energy and infrastructure clients. Their funding and partnership activity suggests they are entering a commercial scaling phase.

Useful evidence for a one-pager:
- London could support investor, utility and infrastructure client engagement.
- The company may prefer flexible space before committing to a large headquarters.
- Risk: Manchester remains the main HQ and talent base.
"""


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

    (NOTES_DIR / "northstar_robotics.md").write_text(northstar_note, encoding="utf-8")
    (NOTES_DIR / "helio_grid.md").write_text(helio_note, encoding="utf-8")

    # Convenience copies for manual GitHub web uploads where everything is uploaded flat.
    shutil.copy2(DB_PATH, BASE / "briefing_demo.db")
    (BASE / "northstar_robotics.md").write_text(northstar_note, encoding="utf-8")
    (BASE / "helio_grid.md").write_text(helio_note, encoding="utf-8")

    print(f"Demo database created at: {DB_PATH}")
    print(f"Flat upload copy created at: {BASE / 'briefing_demo.db'}")
    print(f"Demo notes created at: {NOTES_DIR}")


if __name__ == "__main__":
    main()
