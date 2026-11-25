# Security Policy

Plant Watch reads plant data, so we take security reports seriously.

## Reporting a vulnerability

**Please don't report security problems in a public issue.**

Use one of these instead:

1. **GitHub private reporting:** go to the Security tab of this repository and choose "Report a vulnerability."
2. **Email:** security@humagined.ai

Please include:

- What you found and where (file, skill, or step)
- How to reproduce it, using the synthetic shop in `demo/` if you can
- What you think the impact is

Never send real plant data, credentials, or customer information in a report.

## What to expect

Plant Watch is community supported on a best-effort basis. For security reports we aim to:

- Acknowledge your report within 5 business days
- Tell you whether we can reproduce it and what we plan to do
- Credit you in the release notes if you'd like

We'll ask you to keep the issue private until a fix is released.

## Scope

In scope:

- Anything that sends data off the user's machine without the user's action
- Anything that writes to or changes source files Plant Watch should only read
- Prompt injection through exports, spreadsheets, or report files that changes Plant Watch's behavior
- Exposure of file contents or paths in exports beyond what the user asked for

Out of scope:

- How Claude itself handles data (see Anthropic's policies)

## Supported versions

Only the latest release on the `main` branch receives security fixes.
