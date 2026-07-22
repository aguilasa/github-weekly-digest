# github-weekly-digest

A Claude Cowork / Claude Code plugin that turns "check what happened in
this repo this week" into one request: it fetches pull request and
issue activity from GitHub, writes it into an Excel workbook, and
emails a summary with the workbook attached.

Built for the Claude Cowork course lab (module 12): a plugin chaining
two apps — GitHub and Gmail — into one automated workflow, using their
MCP connectors rather than a standalone script with its own
credentials, matching how a real Cowork plugin operates.

## Plugin structure

```
github-weekly-digest/
├── plugin.json                        # manifest: commands, skills, connectors
├── pyproject.toml / uv.lock            # dependencies, managed by uv
├── .env.example                        # config template (copy to .env)
├── commands/
│   └── digest.md                       # what the /digest command does
├── skills/
│   └── github-weekly-digest/
│       ├── SKILL.md                    # step-by-step MCP-driven instructions
│       └── build_workbook.py           # the one non-MCP step: data -> .xlsx
├── references/
│   ├── email_template.md               # shape of the email/workbook output
│   └── sample-digest-data.json         # fixture for testing build_workbook.py
└── output/                             # generated data/workbooks (gitignored)
```

## How it works

Unlike a cron script that manages its own GitHub token and SMTP
credentials, this plugin is designed to run **inside a live Claude Code
or Cowork session**, using the connectors already authorized there:

1. **GitHub MCP connector** (`mcp__github__search_pull_requests`,
   `mcp__github__search_issues`) fetches open and recently-updated PRs
   and issues for the target repo.
2. Claude writes that data to `output/digest-data.json`.
3. **`build_workbook.py`** (local, no network) turns that JSON into an
   `.xlsx` workbook and prints an email-ready summary. This is the only
   part that isn't an MCP call, because no connector builds Excel
   files — keeping it as a small, pure script means it can be tested
   in isolation (see below) instead of only inside a live session.
4. **Gmail MCP connector** (`mcp__gmail__send_email`) sends the summary
   with the workbook attached.

Full step-by-step instructions Claude follows are in
`skills/github-weekly-digest/SKILL.md`.

## Setup

1. **MCP connectors**: make sure the GitHub and Gmail MCP connectors
   are available in your Claude Code / Cowork session (already
   authorized once you've connected those accounts — no separate token
   or app password to manage for this plugin).

2. **Python environment** (only needed for `build_workbook.py`): this
   project is managed with [`uv`](https://docs.astral.sh/uv/). From the
   project root:

   ```
   uv sync
   ```

3. **Configuration**: copy the example env file and fill it in:

   ```
   cp .env.example .env
   ```

   - `REPO`: the `owner/name` repo to digest.
   - `DAYS`: look-back window in days (default 7).
   - `EMAIL_TO`: recipient address for the digest email.

## Usage

Inside a Claude Code / Cowork session, describe the task in natural
language or use the slash command:

```
/digest aguilasa/css-windify
```

Claude then follows `SKILL.md`: search GitHub via MCP, write
`output/digest-data.json`, run `build_workbook.py`, and send the email
via the Gmail MCP connector.

### Testing the workbook step in isolation

`build_workbook.py` has no network dependency, so it can be verified
without a live MCP session, using the bundled fixture:

```
uv run skills/github-weekly-digest/build_workbook.py references/sample-digest-data.json --out output/sample-test.xlsx
```

This prints the summary text and writes a two-sheet workbook
(`Pull Requests`, `Issues`) to the given path — useful for confirming
the plugin still works correctly if you ever fork or hand it off to a
teammate (the "handover test").

### Example scenario

Running the full pipeline against `aguilasa/css-windify` sends an email
like:

```
Subject: GitHub weekly digest: aguilasa/css-windify

Weekly GitHub digest for aguilasa/css-windify (last 7 days)

Pull requests: 0 open, 0 merged this period
Issues: 0 open, 0 closed this period

Full details attached in the Excel workbook.
```

(That repo has no PRs or issues at all as of this writing — the
workbook's sheets correctly show "No activity in this period" rather
than failing or looking broken. The `references/sample-digest-data.json`
fixture shows what a non-empty digest looks like.)

### Scheduling it

Set this up as a Cowork **Scheduled Task** (see the course's module 9)
by describing the recurring request, e.g. "every Monday at 8am, run the
github weekly digest for aguilasa/css-windify" — Cowork handles the
recurrence; this plugin handles the workflow.

## Error handling

- `build_workbook.py` reports missing/unreadable input files and
  malformed JSON with a clear one-line message and a non-zero exit
  code, instead of a raw traceback.
- Zero PRs or issues in the look-back window renders as an explicit "No
  activity in this period" row, so a quiet week and a broken fetch
  never look the same.
- If the Gmail MCP send step fails, the workbook already written to
  `output/` is left in place so the run doesn't need to re-fetch data
  from GitHub before retrying.

## License

MIT
