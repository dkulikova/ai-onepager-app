# AI Business Briefing Assistant

A Streamlit prototype that turns a free-text company request into an editable PowerPoint one-pager.

## What the tool does

The app lets a user:

1. Type any company name.
2. Select a brief type: `Investment one-pager` or `Company brief`.
3. Add messy notes or upload a `.txt` / `.md` file.
4. Match the company to an internal SQLite company intelligence database.
5. Optionally include the latest five Google News RSS articles.
6. Generate a structured one-page PowerPoint company profile.
7. Review the source pack, latest news and AI review output.

## What changed in this version

The database has been expanded from a simple company table into a richer company intelligence layer:

- `companies` — company profile, mission, HQ, founded year, company type, sector and positioning
- `leadership` — executive leadership and roles
- `products` — products, services and platform components
- `financials` — funding, valuation and commercial signals
- `milestones` — company timeline and notable signals
- `projects` — demo brief types and use-case context

The PowerPoint one-pager now includes:

- Mission / positioning
- Growth direction
- Target market
- Company snapshot
- What they do
- Executive leadership
- Funding / commercial signals
- Latest news / signals
- Risks / things to verify
- Recommended next steps
- Company milestones / signals

## Demo companies

The demo database includes real example companies:

- OpenAI
- Anthropic
- Mistral AI

The data is presentation-ready and intentionally concise. It should be treated as a demo company-intelligence layer, not a live financial database.

## Important limitations

- These companies are private, so full financial performance is not publicly available like it would be for listed companies.
- Funding and valuation details are included as commercial signals, not complete financial statements.
- Leadership, product names, funding and valuation can change quickly.
- Google News RSS results are useful as signals, but can be duplicated, irrelevant, incomplete or behind paywalls.
- The PowerPoint output is an AI-generated first draft and should be reviewed before external use.

## Run locally

```bash
pip install -r requirements.txt
python setup_demo_data.py
streamlit run app.py
```

## Streamlit Cloud deployment

If you upload everything flat to GitHub, include at minimum:

```text
app.py
setup_demo_data.py
requirements.txt
briefing_demo.db
openai.md
anthropic.md
mistral_ai.md
```

If you preserve folders, include:

```text
app.py
setup_demo_data.py
requirements.txt
data/briefing_demo.db
notes/openai.md
notes/anthropic.md
notes/mistral_ai.md
```

Set Streamlit main file path to:

```text
app.py
```

## API keys

The app runs in `Demo fallback` mode without an API key. For real model output, add secrets in Streamlit Cloud:

```toml
OPENAI_API_KEY = "your_key_here"
ANTHROPIC_API_KEY = "your_key_here"
```

## Interview talking point

The visible output is a PowerPoint one-pager. The more important concept is the workflow behind it: the tool checks internal company intelligence, retrieves supporting notes, optionally pulls latest external signals, gives the LLM a controlled source pack, and maps the structured output into an editable business artefact.
