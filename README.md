# Plant Watch

**Find where your plant's numbers disagree, and why.**

Plant Watch is a free, open source plugin for Claude. It reads your plant's exports, spreadsheets and reports, finds the metrics that show up in more than one place, and tells you where they don't match. Then it goes one step further: it shows you whether the numbers disagree because the data is wrong, or because two teams define the metric differently.

Most mismatch tools stop at "these two numbers are different." Plant Watch asks what each number is supposed to mean.

A holiday gift to manufacturers from [Humagined Inc.](https://www.humagined.ai)

---

## What it finds

1. **Numbers that disagree across systems.** OEE in the MES says one thing, the Excel bowler says another, the board report says a third.
2. **Definitions that don't match.** Scrap counted at the machine vs. at final inspection. Downtime with or without planned changeovers. Same name, different metric.
3. **Spreadsheet errors.** Broken formulas, hard-coded overrides, ranges that stopped growing, links to files that moved.
4. **Reports that drifted from the data.** A weekly pack that no longer ties back to the exports it was built from.

The findings land in one **Metric Mismatch Report**, organized by value stream, with each finding traced to the file, page or cell it came from.

## Are you ready for AI agents?

The report opens with a score from 0 to 100. It asks five plain questions of the files it read:

1. Are all of your definitions agreed upon?
2. Do your records connect?
3. Do your data sources hold up?
4. Does someone own each number?
5. Do your reports match your data?

Each check counts the same. One that can't be measured is left out, and the report says so. If a check is under half, the score stays below 65. It isn't a promise about an agent, or about files that weren't read. Data teams call this knowledge graph readiness.

## What it doesn't do

- **It doesn't send your files or results.** Plant Watch has no cloud service, account or telemetry, and its code makes no network calls. It reads the files you point it at and writes its results in a `plant-watch` folder.
- **It doesn't change your source data.** It reads your exports and spreadsheets. It doesn't edit them, and it refuses to save a result over them.
- **It doesn't decide what your metrics mean.** It asks you, once, and records your answer. Your people own the definitions.

Claude itself processes what you share in your conversation under your Claude plan's terms. Plant Watch adds no other destination.

If you give it a SharePoint or OneDrive report link, Claude reads that report through your own Microsoft 365 connection. If you skip links, Plant Watch doesn't use it.

## Try it in 10 minutes with the synthetic shop

You don't need to touch your own data to see what Plant Watch does. The `demo/` folder, which comes in the download, holds a synthetic shop, Northfork Precision: ERP and MES exports, a maintenance log, an Excel shift log and weekly report packs, with real-world style mismatches planted in it.

1. Install the plugin (below).
2. Say: **"Set up Plant Watch with the synthetic shop."**

You'll get a Metric Mismatch Report for Northfork in a few minutes.

## Install

**You need:** the Claude desktop app and a paid Claude plan.

1. **Download** `plant-watch.zip` from [humagined.ai/plantwatch](https://www.humagined.ai/plantwatch). Don't unzip it.
2. **Open the Claude desktop app.** Go to **Customize**, then **Plugins**. Choose **Add**, then **Upload plugin**.
3. **Choose `plant-watch.zip`.** Claude will remind you to install only plugins you trust. That's normal.

Then start a new chat and say **"Set up Plant Watch."**

**Can't upload it?** If your company manages Claude for you, your admin may need to add the plugin. Send them this page. The code is here for them to review.

### For developers (Claude Code)

```
/plugin marketplace add humagined/plant-watch
/plugin install plant-watch@humagined
```

Or download the source and load it locally.

## Using it

| Say this | What happens |
|---|---|
| "Set up Plant Watch" | First run. Finds your exports, spreadsheets and reports, asks what each shared metric should mean, and builds the Metric Mismatch Report with an Excel audit. |
| "Check this report" | Adds a report PDF or link so its figures get checked against your data on each run. |
| "Run Plant Watch" / "What changed?" | Re-checks after new exports land. Shows what's new, what cleared, and what still disagrees. |
| "Export the findings" | Saves findings as CSV or JSON, and your confirmed definitions as a file you keep. |

## Walk through your report

Questions about what you found? [Book a free 30-minute working session](https://www.humagined.ai/book) to walk through your report with us. Available through January 31, 2027.

## Support

Plant Watch is community supported. Humagined responds to issues and discussions on a **best-effort basis**. There is no SLA.

- **Found a bug?** [Open an issue](../../issues/new/choose).
- **Have a question or idea?** Start a [discussion](../../discussions).
- **Found a security problem?** Don't open a public issue. See [SECURITY.md](SECURITY.md).

**Please don't post real plant data in an issue.** Use the synthetic shop, or redact names, part numbers, customers and figures first.

## Contributing

Pull requests are welcome. Keep changes focused, test them against the synthetic shop in `demo/`, and describe what you changed and why. By contributing, you agree your contribution is licensed under Apache 2.0.

## License

Apache License 2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE).

Plant Watch is provided as-is, without warranty. Check its findings before you act on them.

Plant Watch is built for Claude. It is not made, endorsed or supported by Anthropic. Humagined, Plant Watch and TessereX are trademarks of Humagined Inc. Claude is a trademark of Anthropic, PBC.
