---
name: grant-decision-snapshot
description: Use when a program officer needs a short, skimmable summary of a grant request before a review meeting. Turns the key facts of a request into a one-paragraph decision snapshot plus the open questions to resolve. Trigger on "decision snapshot", "summarize this request", "prep me for review", or pasted request details.
---

# Grant Decision Snapshot

## What it does
Takes the key facts of a grant request and produces a one-paragraph snapshot a
reviewer can skim in 30 seconds, plus a short list of what still needs answering.

## When to use it
- Prepping a docket or board packet.
- Walking into a review meeting needing the gist of each request.
- Handing a request to a colleague who needs context fast.

## Safety first
- De-identify applicant and organization names before pasting into a general
  AI tool: "Nonprofit B", "Program C".
- Never paste scoring, committee deliberations, or denial rationale.

## How it works
1. Paste the de-identified request facts as INPUT.
2. The skill returns a snapshot in a fixed structure.
3. You confirm every fact against akoyaGO before it informs a decision.

## The prompt
```
You are summarizing a grant request for a program officer's review.
GOAL: produce a short decision snapshot.
FORMAT:
- Snapshot: one paragraph (who, what, how much, why now)
- Strengths: up to 3 bullets
- Watch-outs: up to 3 bullets
- To confirm: questions to resolve before deciding
CONSTRAINTS: Use only the facts provided. Do not invent amounts, dates, or
outcomes. If something important is missing, put it under "To confirm".
INPUT (de-identified):
<paste request facts here>
```

## Guardrails
- The snapshot summarizes; it does not recommend approve/deny.
- Every number and date is verified in akoyaGO before the meeting.

## Example
INPUT: facts for a $25,000 capacity-building request from a youth nonprofit.
OUTPUT: "Snapshot: Nonprofit B requests $25,000 over 12 months to fund... 
Watch-outs: budget lacks detail on staffing. To confirm: matching funds secured?"
