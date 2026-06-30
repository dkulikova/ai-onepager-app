# AI Business Briefing Assistant

A lightweight Streamlit prototype that turns structured company data, Markdown/Obsidian-style notes, and user-added context into an editable one-page PowerPoint company profile.

## What the app demonstrates

- Structured data access from a local SQLite database
- Unstructured note retrieval from Markdown files
- Automatic latest-news retrieval from Google News RSS for the selected company
- LLM-powered structured drafting
- Review step for unsupported claims and missing information
- Downloadable single-slide `.pptx` output using a preset company-profile template
- Fixed “Latest news / recent signals” section in the one-pager

## Main files

- `app.py` — the Streamlit app and PowerPoint generation logic
- `setup_demo_data.py` — creates the demo SQLite database and sample notes
- `requirements.txt` — Python dependencies for Streamlit Cloud, including `feedparser` for RSS retrieval
- `briefing_demo.db` — flat-upload copy of the demo SQLite database
- `openai.md`, `anthropic.md` and `mistral_ai.md` — flat-upload copies of the demo notes
- `data/briefing_demo.db` — foldered copy of the demo SQLite database
- `notes/` — foldered copy of the demo notes
- `.streamlit/secrets.toml.example` — example secret names for API keys

The app supports both foldered and flat GitHub uploads. If everything is uploaded to the same GitHub folder, it will look for `briefing_demo.db` and the `.md` notes beside `app.py`. If folder structure is preserved, it will use `data/briefing_demo.db` and `notes/`.


## Demo companies

The included demo database contains three real AI-company examples:

- **OpenAI** — frontier AI research and deployment company; demo profile focuses on ChatGPT, Codex, API and enterprise AI platform positioning.
- **Anthropic** — AI safety and research company; demo profile focuses on Claude Enterprise, governance, data controls and reliable AI systems.
- **Mistral AI** — Paris-headquartered AI company; demo profile focuses on open models, enterprise agents and deployment portability.

## Deploying on Streamlit Cloud

1. Upload these files to GitHub:
   - `app.py`
   - `setup_demo_data.py`
   - `requirements.txt`
   - `README.md`
   - `briefing_demo.db`
   - `openai.md`
   - `anthropic.md`
   - `mistral_ai.md`

2. In Streamlit Cloud, create a new app from the GitHub repo.

3. Use this as the main file path:

```text
app.py
```

4. Deploy the app.

The app will work in **Demo fallback** mode without an API key. The RSS feature does not require an API key; it retrieves the latest five Google News RSS results for the selected company and passes them into the source pack.

## Optional API keys

To use a real model, add secrets in Streamlit Cloud app settings:

```toml
OPENAI_API_KEY = "your_key_here"
ANTHROPIC_API_KEY = "your_key_here"
```

Do not commit real API keys to GitHub.

## Interview positioning

A simple way to explain the tool:

> The visible output is a one-page PowerPoint company profile, but the underlying workflow is more than text generation. It retrieves structured database records, combines them with unstructured notes and recent news signals, asks an LLM to create structured fields, runs a review step, and maps the result into an editable business-ready template.

A simple non-technical analogy:

> It works like a fast first-draft analyst. It pulls together the available evidence, organises it into a familiar company-profile format, drafts the slide, and then flags where a human still needs to check the judgement.

## Limitations to mention

- The demo database is small and uses real-company example profiles for OpenAI, Anthropic and Mistral AI; figures are indicative and should be verified before external use.
- Keyword note retrieval is simple; a production version could use embeddings/vector search.
- The AI can overstate weak evidence if the source material is poor.
- RSS/news retrieval can return duplicate, irrelevant, paywalled or weak articles; these should be treated as signals to verify, not definitive evidence.
- The PowerPoint is a first draft and still requires human review.
- A production version would need access permissions, audit logs, source citations and approval workflows.

## Future production architecture

A more robust version could include:

- CRM / SharePoint / Snowflake / SQL Server connectors
- MCP-style integration layer
- RAG with citations
- Approved news/data provider integration instead of a public RSS feed
- structured output validation
- user feedback loop
- human approval before export or client sharing
