---
name: email-thread-summary
description: Use when someone pastes a long or messy email thread and needs to catch up fast. Produces a short summary, the decisions and open questions, and a ready-to-edit reply. Trigger on "summarize this thread", "catch me up", "what is this email asking for", or any pasted multi-message email chain.
---

# Email Thread Summary

## What it does
Turns a long email thread into a quick catch-up — what happened, what is being
asked of us, and a draft reply you can edit and send.

## When to use it
- A thread is too long to read before a meeting.
- You are picking up a conversation someone else started.
- You need a reply but do not want to start from a blank page.

## Safety first
- De-identify before pasting into a general AI tool: use "Grantee A", "Staff B".
- Never paste PII, confidential decisions, or donor/financial specifics.

## How it works
1. Paste the de-identified thread as INPUT.
2. The skill summarizes it into a fixed structure.
3. It drafts a reply; you verify and send it from akoyaGO / Outlook.

## The prompt
```
You are helping a foundation staff member catch up on an email thread.
GOAL: summarize the thread and draft a reply.
FORMAT:
- Summary: 3 bullets
- Asked of us: what we need to do or decide
- Open questions: anything unclear
- Draft reply: short, warm, professional
CONSTRAINTS: Do not invent facts. If key info is missing, list up to 3
questions instead of guessing.
INPUT (de-identified):
<paste thread here>
```

## Guardrails
- Facts come from the thread only — nothing added.
- Verify names, dates, and amounts in akoyaGO before sending.

## Example
INPUT: an 8-message thread about a late grant report.
OUTPUT: "Summary: grantee missed the May report, requested a 2-week extension,
approval pending... Draft reply: Hi [name], thanks for letting us know..."
