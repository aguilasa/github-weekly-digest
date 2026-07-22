# Skill: github-weekly-digest

## What this skill does

Given a GitHub repository, produces a weekly activity digest: open and
recently-updated pull requests and issues, written to an Excel workbook,
summarized and emailed to a configured recipient.

## Connectors it uses

- **GitHub MCP connector** (`mcp__github__search_pull_requests`,
  `mcp__github__search_issues`) — fetches the data. No token or `gh`
  CLI setup needed beyond the MCP connector already being available in
  the Cowork / Claude Code session.
- **Gmail MCP connector** (`mcp__gmail__send_email`) — sends the
  summary email with the workbook attached. No SMTP credentials or App
  Password needed; it uses the Gmail account already authorized for the
  MCP connector.
- **`build_workbook.py`** (local script, no network) — the one part
  no connector can do: turning fetched data into an `.xlsx` file. Kept
  as a small, pure, independently-testable script (see
  `references/sample-digest-data.json` for a fixture that exercises it
  without a live session).

## How to carry it out

When this skill is invoked for a target `owner/repo` (and a look-back
window in days, default 7):

1. **Fetch pull requests** with the GitHub MCP connector:
   - `mcp__github__search_pull_requests` with query
     `repo:{owner}/{repo} is:pr is:open` (all currently open PRs).
   - `mcp__github__search_pull_requests` with query
     `repo:{owner}/{repo} is:pr updated:>={since_date}` (PRs merged or
     otherwise updated in the window).
   - Merge the two result sets, de-duplicating by PR number.

2. **Fetch issues** the same way with
   `mcp__github__search_issues`, using `is:issue` instead of `is:pr`.

3. **Write a `digest-data.json` file** (e.g. under `output/`) shaped as:

   ```json
   {
     "repo": "owner/repo",
     "days": 7,
     "pull_requests": [
       {"number": 1, "title": "...", "author": "...", "state": "open",
        "created_at": "...", "updated_at": "...", "merged_at": null,
        "url": "..."}
     ],
     "issues": [
       {"number": 2, "title": "...", "author": "...", "state": "closed",
        "created_at": "...", "updated_at": "...", "closed_at": "...",
        "url": "..."}
     ]
   }
   ```

   Map each GitHub MCP result's `number`, `title`, `user.login` (as
   `author`), `state`, `created_at`, `updated_at`,
   `merged_at`/`closed_at`, and `html_url` (as `url`) into this shape.

4. **Run the local script** to build the workbook and get the email
   text:

   ```
   uv run skills/github-weekly-digest/build_workbook.py output/digest-data.json
   ```

   It prints the email summary to stdout and a final
   `WORKBOOK_PATH=...` line with the generated `.xlsx` path.

5. **Send the email** with the Gmail MCP connector
   (`mcp__gmail__send_email`), using the printed summary as the body,
   `GitHub weekly digest: {owner}/{repo}` as the subject, and the
   workbook path from step 4 in `attachments`.

## Failure handling

- `build_workbook.py` reports a clear one-line error (missing/invalid
  JSON input, missing `repo` field) instead of a raw traceback, and
  exits non-zero.
- If a GitHub MCP search returns zero results, still write an empty
  `pull_requests` / `issues` list — `build_workbook.py` renders that as
  a "No activity in this period" row rather than failing or producing a
  blank sheet.
- If the Gmail MCP send fails, the workbook that was already generated
  is left in place, so the run can be retried without re-fetching data.

## Reused by

`commands/digest.md` — the `/digest` command wraps this skill for
manual or scheduled invocation.
