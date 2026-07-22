# Skill: github-weekly-digest

## What this skill does

Given a GitHub repository, produces a weekly activity digest: open and
recently-updated pull requests and issues, written to an Excel workbook,
summarized and emailed to a configured recipient.

## Connections it needs

- **GitHub**: the `gh` CLI, authenticated on the machine running the
  script (`gh auth login`). No separate API token is managed by this
  skill.
- **Email**: a Gmail account with an App Password (see the project
  README for how to generate one), used over SMTP to send the summary.

## How to carry it out

Run `script.py` in this folder (via `uv run`, from the project root, so
it can find `.env` and write to `output/`):

```
uv run skills/github-weekly-digest/script.py --repo owner/name [--days 7] [--dry-run]
```

- `--repo` is required unless `REPO` is set in `.env`.
- `--days` controls the look-back window for "recently updated" items
  (default 7).
- `--dry-run` builds the workbook and prints the summary but does not
  send an email — use this to verify data retrieval before wiring up
  email credentials.

## Failure handling

The script distinguishes between expected, reportable failures (missing
`gh` auth, missing email config, SMTP errors) and lets those print a
clear one-line error instead of a raw traceback. A failure to send email
does not delete the workbook that was already generated.

## Reused by

`commands/digest.md` — the `/digest` command wraps this skill for
manual or scheduled invocation.
