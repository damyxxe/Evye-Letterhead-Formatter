# Changelog

## v1.2.0 — 2026-07-22

**Formatting parity & audit remediation.** Closed every gap in the Notion→Word formatting
audit; every Notion block type now maps to a Word representation, nothing is silently
dropped.

### Renderer (`word_generator.py`)

- **Run-level formatting**: bold/italic/underline/strike/code/hyperlinks now render as real
  Word run formatting instead of one plain run per paragraph. New run schema
  `{"t","b","i","u","s","c","link"}`, threaded through `_add_para`, Notion rich-text
  extraction (`_extract_notion_runs`), and markdown inline parsing (`_parse_inline`).
- **New block types**: `numbered` (auto-numbered, nesting-aware), `todo` (☑/☐), `code_block`
  (monospace), `image` (downloaded + embedded, ≤133mm, degrades to visible placeholder on
  failure — never silent), `generic_table` (N-column fallback for 1-col or 4+-col tables —
  nothing truncated).
- **Notion conversion rewrite**: full block vocabulary (`to_do`, `toggle`, `code`, `image`,
  `bookmark`/`embed`, `file`/`pdf`/`video`/`audio`, `equation`, `child_page`, and more), Notion
  heading re-levelling (H1/H2/H3 → Word H2/H3/H4, all three now distinct), and a
  **no-silent-loss fallback**: any unrecognized block renders as visible
  `[Unsupported block: type]` text.
- **Markdown path parity**: same inline formatting, numbered lists, todos, and fenced code
  blocks as the Notion path.
- **Test suite**: 36 pytest tests, including `tests/test_integration_notion.py` — a
  full-fixture integration test asserting every supported block type renders correctly, the
  permanent regression gate for the coverage matrix.
- Version identity: `word_generator.__version__` (now `"1.2.0"`), surfaced via
  `GET /docs/api/health`.

### Repo hygiene

- Removed dead WeasyPrint/PDF-era templates and one-off migration scripts (`fix_*.py`,
  `generator.py`, `templates/{base,contract,letter,quotation}.html`).
- Rewrote `README.md`; fixed an INSTALL.txt font-extension mismatch (`.otf` → `.ttf`).

### n8n workflow (`Evye Assistant`, `KmAFj9dA7wCFGxnQ`)

- Removed 6 dead/redundant nodes (`Extract Request`, `GET Notion Blocks`,
  `Extract Table IDs`, `Has Tables?`, `GET Table Rows`, `Rebuild Blocks`) — 46 → 40 nodes.
- **Fixed nested-table data loss**: the old table-discovery pass only scanned top-level
  blocks, so any table nested inside a toggle/column/list silently vanished. The rewritten
  `Fetch Block Children` node now attaches `_rows` during the same recursive walk that
  discovers blocks at every depth.
- Removed a redundant first-pass Notion fetch (`GET Notion Blocks`) whose output was
  entirely discarded downstream.
- Removed a vestigial dummy-LLM node.
- Hardcoded Notion/Telegram tokens moved to `$env.NOTION_TOKEN_EVE`/`$env.TELEGRAM_TOKEN_EVE`
  (wired via `/Users/alfred/docker/docker-compose.override.yml` + `.env` on Alfred, same
  pattern as Media Monitoring's secrets). **Note: relocated, not rotated** — same token
  values. Rotation deferred (see Pending below).
- `/help` text now reports the version.

### Deployed

- `evye-docgen` on Alfred (`co.evye.docgen`, port 8091) — verified live at v1.2.0.
- `Evye Assistant` n8n workflow — updated, cycled (deactivate→activate to re-register the
  Telegram webhook), confirmed active with zero hardcoded tokens remaining.

### Baseline / rollback

- Tags: `eve-v1.1.0` (pre-v1.2 baseline — the deployed-but-uncommitted code as of the start
  of this work), `eve-v1.2.0` (this release).
- n8n workflow backups: `~/ClaudeCode/Alfred Setup/backups/evye-assistant-pre-v1.2-2026-07-22.json`
  (pre-surgery, 46 nodes), `evye-assistant-v1.2-final-2026-07-22.json` (post-surgery, 40
  nodes, active).
- Rollback renderer: `git checkout eve-v1.1.0 -- word_generator.py api.py` + redeploy.
- Rollback n8n workflow: PUT the pre-surgery backup (n8n's public API rejects extra
  `settings` keys like `callerPolicy`/`availableInMCP`/`binaryMode` with "must NOT have
  additional properties" — strip to `{saveExecutionProgress, saveManualExecutions,
  saveDataErrorExecution, saveDataSuccessExecution, executionTimeout, errorWorkflow,
  timezone, executionOrder}` first), then deactivate→activate.

### Pending (not part of this release)

- **Secret rotation** — the Notion integration token and Telegram bot token were hardcoded
  in n8n workflow code prior to this release and passed through a Claude session transcript
  during the build (inspecting old node code to design the replacement). They're now in env
  vars but are the *same, unrotated* values. Explicitly deferred by Michael on 2026-07-22
  pending an urgent live document. Do not treat as resolved.
- Live UAT via Telegram with a real multi-block document.
- Out of scope for v1.3+: run-formatting inside table cells, Notion color annotations →
  Word colors, callout icon/tint rendering, LLM-based intent extraction.
