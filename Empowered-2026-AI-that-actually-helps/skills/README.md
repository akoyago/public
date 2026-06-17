# akoyaGO AI Skills — starter set

Five simple, reusable **AI skills** for foundation staff. Each one is a single
`SKILL.md` file: a saved set of instructions that teaches an AI to do one job
well, the same way every time.

| Skill | Job it does |
|-------|-------------|
| `email-thread-summary` | Catch up on a long thread + draft a reply |
| `tone-rescue` | Rewrite a blunt draft into neutral + warm versions |
| `grant-decision-snapshot` | Turn request facts into a skimmable review snapshot |
| `scholarship-letters` | Generate award / not-selected / incomplete templates |
| `call-to-crm` | Turn call notes into an email, tasks, and a timeline note |

## Shared anatomy

Every skill follows the same simple shape, which makes them easy to read, teach,
and write your own:

1. **Frontmatter** (`name` + `description`) — the label. The `description` is the
   most important line: it tells the AI *when* to reach for this skill.
2. **What it does** — one plain sentence.
3. **When to use it** — the situations it fits.
4. **Safety first** — de-identify; never paste the 3 Red Lines.
5. **How it works** — the steps.
6. **The prompt** — the reusable prompt, built on ROLE · AUDIENCE · GOAL · INPUT ·
   FORMAT · CONSTRAINTS.
7. **Guardrails** — "do not invent facts"; verify against akoyaGO.
8. **Example** — a tiny before/after so the result is obvious.

## How to use one
Copy the prompt block, paste your **de-identified** input where it says, and edit
the result. The record in akoyaGO / Dynamics / Business Central is always the
source of truth.
