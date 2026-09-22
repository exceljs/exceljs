# Security Policy

`exceljs-hardened` is an unofficial security fork of
[exceljs](https://github.com/exceljs/exceljs), maintained specifically
because the upstream project currently has no working way to receive
vulnerability reports. This document describes how to report issues
**in this fork**.

## Supported versions

| Version | Supported |
|---|---|
| 4.4.x (latest) | ✅ |
| < 4.4.1 | ❌ |

Only the latest published version receives security fixes. There is no
long-term support branch — this is a small volunteer effort, not a
funded security team.

## Reporting a vulnerability

**Please use [GitHub Security Advisories](https://docs.github.com/en/code-security/security-advisories/guidance-on-reporting-and-writing/privately-reporting-a-security-vulnerability)
for this repository — private reporting is enabled.**

1. Go to the **Security** tab of this repository
2. Click **Report a vulnerability**
3. Fill in the advisory draft with as much detail as you can:
   - Affected version(s)
   - A clear description of the vulnerability and its impact
   - Steps to reproduce, or ideally a minimal proof of concept
   - Suggested fix, if you have one

This creates a private draft only you and the maintainer can see, so it's
safe to include a working exploit.

**Do not open a public GitHub issue for security reports.** Anything
filed as a regular issue will be treated as public from the moment it's
posted.

## What to expect

- **Acknowledgement:** best-effort within 7 days.
- **Triage:** I'll confirm whether it's in scope, ask follow-up
  questions if needed, and give a rough timeline.
- **Fix & disclosure:** once a patch is ready, I'll coordinate a release
  date with you. Absent a response from you within 90 days, I may
  publish the advisory and fix on my own timeline — this mirrors
  standard industry practice and exists to make sure known issues don't
  sit unpublished indefinitely.
- **Credit:** reporters are credited in the published advisory and
  release notes unless you ask to stay anonymous.

## Scope

**In scope:** vulnerabilities in this fork's own code, including the
patches listed in [`README.md`](./README.md) and any code inherited
unchanged from upstream `exceljs@4.4.0`.

**Out of scope:**
- Vulnerabilities that require the *host application* to misuse the
  library in ways the documentation already warns against (e.g. passing
  fully-qualified, attacker-controlled absolute paths directly into
  `addImage()` — see the note in
  [`lib/utils/media-path-guard.js`](./lib/utils/media-path-guard.js)).
  These are still worth reporting so we can improve the docs or add
  further guardrails, just categorized differently.
- Denial-of-service reports that only reproduce with resource limits
  already disabled by the caller (e.g. `maxEntryUncompressedSize` raised
  to an unreasonable value on purpose).
- Anything specific to upstream `exceljs` that has already been reported
  there — please also report it upstream if any of their channels are
  working for you; this fork can only fix what it knows about.

## Why this policy exists

This project exists because upstream `exceljs` currently has:
- No responsive Discord/chat channel
- Private vulnerability reporting disabled on the official repository
- No `SECURITY.md` of its own

If you're a maintainer of the original `exceljs` project and want help
coordinating fixes back upstream, please reach out via a Security
Advisory here — I'd genuinely rather this fork become unnecessary.
