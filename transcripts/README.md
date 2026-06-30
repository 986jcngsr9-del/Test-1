# Call Transcripts

A version-controlled store for meeting/call transcripts, organized by type and
tagged with structured metadata so they can be filtered, searched, and reported on.

Transcripts are pulled from [Granola](https://granola.ai) and committed here as
Markdown files with a YAML frontmatter header for classification.

## Organization

Each call lives in a folder for its **type**:

| Folder | What goes here |
|---|---|
| `office-hours/` | Recurring open/drop-in sessions (e.g. EMEA Office Hours, AI SolCon Office Hours) |
| `team-calls/` | Internal team syncs and cadence calls not tied to one account |
| `account-specific/` | Calls with or about a single customer/prospect (one subfolder per account) |
| `enablement/` | Training, academies, onboarding, internal enablement |
| `deal-reviews/` | RFP reviews, go/no-go calls, deal-specific working sessions |

For `account-specific/`, group by account: `account-specific/<account-slug>/<file>.md`.

## File naming

```
<YYYY-MM-DD>-<short-slug>.md
```

e.g. `2026-06-29-decisioning-studio-sales-team-meeting.md`

## Metadata (frontmatter)

Every transcript starts with a YAML block:

```yaml
---
title: Decisioning Studio Sales Team Meeting
date: 2026-06-29T19:30:00Z      # ISO 8601, UTC
type: team-call                  # office-hours | team-call | account-specific | enablement | deal-review
region: global                   # emea | amer | apac | global
product_area: Decisioning Studio
account: null                    # company name, or null for internal calls
source: granola
source_meeting_id: c9657e08-83e9-4ce4-9cff-350a790c4a5f
attendee_count: 17
tags: [pipeline, collateral, multi-workspace]
---
```

After the frontmatter, the body has two sections: a **Summary** and the
**Full transcript** (verbatim).

## Adding a transcript

Ask Claude to "pull <meeting> from Granola and file it" — Claude finds the
meeting, classifies it, writes the file into the right folder with frontmatter,
and commits.
