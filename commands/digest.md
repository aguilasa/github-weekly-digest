# /digest

Runs the GitHub weekly digest pipeline for one repository: pulls PR and
issue activity via the GitHub MCP connector, builds an Excel workbook,
and emails a summary via the Gmail MCP connector.

## What it does, step by step

See `skills/github-weekly-digest/SKILL.md` for the full instructions.
Summary:

1. Fetch open + recently-updated pull requests and issues for the
   target repo using `mcp__github__search_pull_requests` /
   `mcp__github__search_issues`.
2. Save the results as `output/digest-data.json`.
3. Run `uv run skills/github-weekly-digest/build_workbook.py output/digest-data.json`
   to produce the `.xlsx` workbook and the email summary text.
4. Send the summary + workbook attachment to the configured recipient
   with `mcp__gmail__send_email`.

## When to use it

Run this whenever you want a snapshot of what changed in a repo over
the last week — either manually, or on a recurring Cowork scheduled
task (see the README) to replace a recurring manual status check.

## How to invoke

- From a Claude Code / Cowork session: describe the request in natural
  language, e.g. "run the github weekly digest for aguilasa/css-windify
  and email it to me", or type `/digest aguilasa/css-windify`.
- The GitHub and Gmail MCP connectors must already be available in the
  session; this command does not manage its own credentials.
