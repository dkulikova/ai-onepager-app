# AI Business Briefing Assistant

A lightweight Streamlit prototype that turns structured company data, Markdown/Obsidian-style notes, and user-added context into an editable one-page PowerPoint company profile.

## What the app demonstrates

- Structured data access from a local SQLite database
- Unstructured note retrieval from Markdown files
- LLM-powered structured drafting
- Review step for unsupported claims and missing information
- Downloadable single-slide `.pptx` output using a preset company-profile template

## Main files

- `app.py` — the Streamlit app and PowerPoint generation logic
- `setup_demo_data.py` — creates the demo SQLite database and sample notes
- `requirements.txt` — Python dependencies for Streamlit Cloud
- `briefing_demo.db` — flat-upload copy of the demo SQLite database
- `northstar_robotics.md` and `helio_grid.md` — flat-upload copies of the demo notes
- `data/briefing_demo.db` — foldered copy of the demo SQLite database
- `notes/` — foldered copy of the demo notes
- `.streamlit/secrets.toml.example` — example secret names for API keys

The app supports both foldered and flat GitHub uploads. If everything is uploaded to the same GitHub folder, it will look for `briefing_demo.db` and the `.md` notes beside `app.py`. If folder structure is preserved, it will use `data/briefing_demo.db` and `notes/`.

## Deploying on Streamlit Cloud

1. Upload these files to GitHub:
   - `app.py`
   - `setup_demo_data.py`
   - `requirements.txt`
   - `README.md`
   - `briefing_demo.db`
   - `northstar_robotics.md`
   - `helio_grid.md`

2. In Streamlit Cloud, create a new app from the GitHub repo.

3. Use this as the main file path:

```text
app.py
```

4. Deploy the app.

The app will work in **Demo fallback** mode without an API key.

## Optional API keys

To use a real model, add secrets in Streamlit Cloud app settings:

```toml
OPENAI_API_KEY = "your_key_here"
ANTHROPIC_API_KEY = "your_key_here"
```

Do not commit real API keys to GitHub.

## Interview positioning

A simple way to explain the tool:

> The visible output is a one-page PowerPoint company profile, but the underlying workflow is more than text generation. It retrieves structured database records, combines them with unstructured notes, asks an LLM to create structured fields, runs a review step, and maps the result into an editable business-ready template.

A simple non-technical analogy:

> It works like a fast first-draft analyst. It pulls together the available evidence, organises it into a familiar company-profile format, drafts the slide, and then flags where a human still needs to check the judgement.

## Limitations to mention

- The demo database is small and uses mock data.
- Keyword note retrieval is simple; a production version could use embeddings/vector search.
- The AI can overstate weak evidence if the source material is poor.
- The PowerPoint is a first draft and still requires human review.
- A production version would need access permissions, audit logs, source citations and approval workflows.

## Future production architecture

A more robust version could include:

- CRM / SharePoint / Snowflake / SQL Server connectors
- MCP-style integration layer
- RAG with citations
- structured output validation
- user feedback loop
- human approval before export or client sharing
