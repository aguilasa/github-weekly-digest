# github-weekly-digest

A Claude Cowork / Claude Code plugin that turns "check what happened in
this repo this week" into one command: it fetches pull request and issue
activity from GitHub, writes it into an Excel workbook, and emails a
summary with the workbook attached.

Built for the Claude Cowork course lab (module 12): a plugin chaining
two apps (GitHub and Gmail) into one automated workflow.

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
│       ├── SKILL.md                    # skill description
│       └── script.py                   # the actual pipeline
├── references/
│   └── email_template.md               # shape of the email/workbook output
└── output/                             # generated workbooks (gitignored)
```

## Setup

1. **GitHub CLI**: make sure `gh` is installed and authenticated:

   ```
   gh auth login
   gh auth status
   ```

2. **Python environment**: this project is managed with
   [`uv`](https://docs.astral.sh/uv/). From the project root:

   ```
   uv sync
   ```

   This creates `.venv` and installs `openpyxl` from `uv.lock`.

3. **Configuration**: copy the example env file and fill it in:

   ```
   cp .env.example .env
   ```

   - `REPO`: the `owner/name` repo to digest.
   - `DAYS`: look-back window in days (default 7).
   - `EMAIL_FROM` / `EMAIL_TO`: sender and recipient addresses.
   - `EMAIL_APP_PASSWORD`: a Gmail **App Password**, not your normal
     password. To generate one:
     1. Turn on 2-Step Verification on the Gmail account, if not already on.
     2. Go to <https://myaccount.google.com/apppasswords>.
     3. Create an app password (any name, e.g. "github-weekly-digest").
     4. Paste the 16-character password into `.env`.

## Usage

Dry run (fetches data, builds the workbook, skips the email — good for
a first test):

```
uv run skills/github-weekly-digest/script.py --repo aguilasa/css-windify --dry-run
```

Full run (also sends the email):

```
uv run skills/github-weekly-digest/script.py --repo aguilasa/css-windify
```

Both `--repo` and `--days` are optional if `REPO` / `DAYS` are set in
`.env`.

From inside a Claude Code or Cowork session, this can also be triggered
by describing the task in natural language ("run the github weekly
digest for aguilasa/css-windify") — see `commands/digest.md`.

### Example scenario

Running the command against `aguilasa/css-windify` produces
`output/digest-aguilasa-css-windify-2026-07-22.xlsx` with two sheets
(`Pull Requests`, `Issues`) and sends an email like:

```
Subject: GitHub weekly digest: aguilasa/css-windify

Weekly GitHub digest for aguilasa/css-windify (last 7 days)

Pull requests: 2 open, 1 merged this period
Issues: 3 open, 0 closed this period

Full details attached in the Excel workbook.
```

### Scheduling it

To replicate Cowork's "Scheduled Tasks" behavior with a plain cron job
(runs every Monday at 8am):

```
0 8 * * 1 cd /path/to/github-weekly-digest && uv run skills/github-weekly-digest/script.py
```

## Error handling

- Missing/unauthenticated `gh` CLI: reported with a one-line fix
  instruction, exits with status 1.
- Missing `--repo` / `REPO`: reported before any network calls, exits
  with status 2.
- No PRs or issues in the look-back window: the workbook records "No
  activity in this period" instead of failing or shipping an empty file.
- Missing email settings, or an SMTP failure: reported clearly; the
  workbook that was already generated in `output/` is kept either way.

## License

MIT
