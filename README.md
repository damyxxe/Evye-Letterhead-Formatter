# Evye Letterhead Formatter ("Eve")

Generates branded Evye LLP Word documents (.docx) from Notion pages or markdown.

- `word_generator.py` — core renderer (python-docx, clones `templates/evye-master.docx`)
- `api.py` — FastAPI wrapper. `POST /docs/api/generate` (markdown), `POST /docs/api/generate-from-blocks` (raw Notion blocks), `GET /docs/api/health`
- Telegram front-end: n8n workflow "Evye Assistant" (`KmAFj9dA7wCFGxnQ`) on Alfred
- Deployed at `alfred:/opt/evye-docgen/` (port 8091, launchd)

## Dev

    pip install -r requirements.txt
    python3 -m pytest tests/ -v

## Deploy (see docs/plans/2026-07-22-eve-v1.2-formatting-parity.md Task 9)

    scp word_generator.py api.py alfred:/opt/evye-docgen/
    # then restart the service and check /docs/api/health reports the new version
