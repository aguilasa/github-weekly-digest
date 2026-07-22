# /digest

Runs the GitHub weekly digest pipeline for one repository: pulls PR and
issue activity, builds an Excel workbook, and emails a summary.

## What it does, step by step

1. Reads `REPO`, `DAYS`, `EMAIL_FROM`, `EMAIL_APP_PASSWORD`, and `EMAIL_TO`
   from the project's `.env` file (or accepts `--repo` / `--days` as
   overrides).
2. Confirms the `gh` CLI is installed and authenticated.
3. Fetches open pull requests, plus PRs updated in the last `DAYS` days,
   for the target repo.
4. Fetches open issues, plus issues updated in the last `DAYS` days, for
   the target repo.
5. Writes both sets into an Excel workbook (`Pull Requests` and `Issues`
   sheets) under `output/digest-<repo>-<date>.xlsx`.
6. Sends an email to `EMAIL_TO` with a short text summary and the
   workbook attached, unless run with `--dry-run`.

## When to use it

Run this whenever you want a snapshot of what changed in a repo over the
last week — either manually, or on a schedule (see the README's cron
example) to replace a recurring manual status check.

## How to invoke

- Manually: `uv run skills/github-weekly-digest/script.py --repo <owner/repo>`
- Dry run (no email sent): add `--dry-run`
- From a Claude Code / Cowork session: describe the request in natural
  language ("run the github weekly digest for aguilasa/css-windify") and
  this skill will be used to carry it out.
