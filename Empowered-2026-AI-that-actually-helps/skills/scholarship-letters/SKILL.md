---
name: scholarship-letters
description: Use when you need to send scholarship decision letters and want consistent, warm wording across award, not-selected, and incomplete outcomes. Generates all three templates with placeholders to fill in. Trigger on "scholarship letter", "award letter", "rejection letter", "not selected", or "decision letters".
---

# Scholarship Letters

## What it does
Generates three reusable letter templates — Award, Not Selected, and Incomplete
Application — written in clear, kind, plain language with placeholders you fill in.

## When to use it
- Decision season, when many letters go out at once.
- You want every applicant to get the same respectful tone.
- You need a starting point that avoids cold or legalistic wording.

## Safety first
- Work in templates with placeholders ([Name], [Program]) — not real applicant
  data — when using a general AI tool.
- Never paste scores, reviewer comments, or denial rationale.

## How it works
1. Tell the skill your program name and tone.
2. It returns three placeholder templates.
3. You merge real names from akoyaGO and review before sending.

## The prompt
```
You are drafting scholarship decision letters for a foundation.
GOAL: produce three short letter templates with [placeholders].
FORMAT:
- Award: congratulations, amount placeholder, next steps
- Not Selected: respectful, encouraging, no reasons given
- Incomplete: what is missing and how to fix it, with a deadline placeholder
CONSTRAINTS: Plain language, warm and professional, no legalese. Do not state
specific reasons for non-selection. Do not invent program details.
DETAILS: program = [name], tone = [warm/formal], signer = [name/title]
```

## Guardrails
- "Not Selected" letters never give individualized reasons.
- Check equity and tone carefully — these letters carry real weight.
- Verify amounts, deadlines, and names in akoyaGO before merging.

## Example
INPUT: program = "Community Futures Scholarship", tone = warm.
OUTPUT: three templates ready to merge, each under 150 words.
