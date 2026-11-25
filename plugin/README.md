# Plant Watch

Find where your plant's numbers disagree, and why. Plant Watch reads a plant's exports, spreadsheets and reports, finds the metrics that show up in more than one place, and builds one Metric Mismatch Report with a readiness for AI agents score.

- **It doesn't send your files or results.** It has no cloud service, account or telemetry, and its code makes no network calls. Claude processes what you share under your Claude plan's terms. Report links are optional, and are read through your own Microsoft 365 connection.
- **It doesn't change your source data.** Results go in a `plant-watch` folder.
- **It doesn't decide what your metrics mean.** It asks you and records your answer.

In Claude Code, install it with:

```
/plugin marketplace add Humagined/plant-watch
/plugin install plant-watch@humagined
```

Say "Set up Plant Watch" to start. Source, tests and a synthetic shop to try it on: https://github.com/Humagined/plant-watch

Community supported, best effort. Apache License 2.0. Plant Watch is built for Claude. It is not made, endorsed or supported by Anthropic. Claude is a trademark of Anthropic, PBC.
