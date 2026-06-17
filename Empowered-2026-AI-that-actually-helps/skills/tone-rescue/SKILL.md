---
name: tone-rescue
description: Use when a draft message is blunt, rushed, or too formal and needs to sound clear and human before it goes out. Returns two rewrites — neutral and warm — without changing the facts. Trigger on "fix the tone", "make this nicer", "soften this", "does this sound okay", or a pasted draft email/letter.
---

# Tone Rescue

## What it does
Rewrites a draft so it reads clearly and kindly, in plain language, and gives
you two tone options to choose from. It changes the wording, never the facts.

## When to use it
- A reply came out sharper than you meant.
- A letter sounds stiff or legalistic.
- You want a second, warmer option before sending.

## Safety first
- De-identify if the draft names real people or organizations.
- Never paste PII, confidential decisions, or donor/financial specifics.

## How it works
1. Paste your draft as TEXT.
2. The skill returns two rewrites (neutral and warm).
3. You pick one, tweak it, and send.

## The prompt
```
You are an editor improving tone for a foundation's outgoing message.
GOAL: rewrite the text below so it is clear, plain, and respectful.
FORMAT: return two versions — (1) Neutral, (2) Warm. Keep each under [X] words.
CONSTRAINTS: Do not change any facts, numbers, or commitments. Avoid legalistic
or robotic phrasing. If a fact seems wrong or missing, flag it — do not invent.
TEXT:
<paste your draft here>
```

## Guardrails
- Only the wording changes; facts and commitments stay exactly as written.
- Read both versions before sending — pick the one that fits the relationship.

## Example
INPUT: "Your report is late. Submit it now."
OUTPUT (Warm): "Hi [name] — a quick nudge that your report is due. Let us know
if you need a few extra days; happy to help."
