# Email summary format

The email body sent by `script.py` follows this shape:

```
Weekly GitHub digest for <owner/repo> (last <N> days)

Pull requests: <open count> open, <merged count> merged this period
Issues: <open count> open, <closed count> closed this period

Full details attached in the Excel workbook.
```

The subject line is `GitHub weekly digest: <owner/repo>`.

The attached workbook (`digest-<repo>-<date>.xlsx`) has two sheets:

- **Pull Requests**: number, title, author, state, created/updated/merged
  timestamps, URL — for every open PR plus any PR updated within the
  look-back window.
- **Issues**: same shape, for issues instead of PRs.

If a sheet would otherwise be empty, it contains a single row reading
"No activity in this period" instead of being left blank, so recipients
can tell the difference between "nothing happened" and "this broke".
