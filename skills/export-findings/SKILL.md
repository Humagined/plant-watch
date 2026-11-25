---
name: export-findings
description: Exports Plant Watch findings as a CSV or JSON file the person can share or load into another tool, and the plant's confirmed definitions as a file they keep. Use when someone says "export the findings", "give me a spreadsheet of the issues", "send this to IT", "I need the findings as a file", "share the mismatch list", "export our definitions", or "export the dictionary".
---

# Export findings

Run `check-dictionary --state plant-watch/state` first and fix what it reports with the person.
Then run `python3 "<this skill's folder>/../../scripts/watch.py" export --state plant-watch/state
--out plant-watch/<file>`, with `.csv` for a spreadsheet (default) or `.json` for another tool
(schema `plant-watch.findings/v1`). When there are confirmed definitions, the same command saves
them beside it as `plant-watch-definitions.json`. Share both files with the person.

Say what's in them in a few lines, from the command's output: how many findings and how many are
open; how many confirmed definitions, and how many figures aren't settled yet (those stay out of
the file). Say the files hold the plant's own data (sources, sheet and cell references, gaps,
owners' names), so they should share them the way they'd share the underlying files.

The export stops there. It writes files in the person's folder and sends nothing. Never upload,
email or post them, and never call another service with them, as part of this skill.

If the command refuses a file name (it exists, or it is among the plant's own files), choose a new
name. Never delete or move the existing file to make room.
