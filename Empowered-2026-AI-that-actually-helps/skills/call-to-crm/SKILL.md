---
name: call-to-crm
description: Use after a donor or grantee call to turn rough notes into a follow-up email, a task list, and a clean timeline note for akoyaGO. Trigger on "log this call", "call notes", "follow up after the call", "turn this into tasks", or pasted meeting/phone notes.
---

# Call to CRM

## What it does
Takes messy call notes and produces three things at once: a short follow-up
email, an owner-and-due-date task list, and a tidy paragraph for the akoyaGO
timeline.

## When to use it
- Right after a donor or grantee call, while it is fresh.
- When notes need to become action, not just sit in a notebook.
- To keep the record consistent across the team.

## Safety first
- De-identify before pasting into a general AI tool: "Donor C", "Officer D".
- Never paste private donor notes, banking details, or restricted-gift specifics.

## How it works
1. Paste your de-identified notes as NOTES.
2. The skill returns an email, a task list, and a timeline note.
3. You verify and save the record in akoyaGO.

## The prompt
```
You are helping log a donor/grantee call for a foundation CRM.
GOAL: turn the notes into follow-up actions.
FORMAT:
- Follow-up email: short, warm, ready to edit
- Tasks: each with an owner placeholder and a due-date placeholder
- Timeline note: 2-3 sentences, factual, for the akoyaGO record
CONSTRAINTS: Use only what is in the notes. Do not exaggerate outcomes or
commitments. If a commitment is unclear, flag it instead of assuming.
NOTES (de-identified):
<paste your notes here>
```

## Guardrails
- No exaggeration — stewardship language stays true to what was said.
- Confirm any promised amounts or dates in akoyaGO before acting.

## Example
INPUT: "Called Donor C, excited about youth program, maybe a fall gift, wants
the impact report."
OUTPUT: email thanking them + tasks (send impact report, schedule fall check-in)
+ timeline note.
