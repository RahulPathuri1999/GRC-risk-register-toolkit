---
name: risk-register
description: "Create, update, and maintain a GRC risk register — a spreadsheet tracking organizational risks with Asset ID/CIA context, Likelihood/Impact scoring, and residual risk after treatment. Use this skill whenever the user asks to create a risk register, add a risk, add a new asset, update likelihood/impact/status/owner on a risk, or asks about risk scoring, inherent vs residual risk, or risk treatment plans. Also use when the user mentions tracking risks for an audit, a security review, or compliance work."
license: Personal project skill — not affiliated with any employer or framework body
---

# Risk Register

Helps produce and maintain a GRC risk register — a live spreadsheet, not a static document —
so risks only need to be entered once and the scoring stays correct automatically as data changes.

| File | Use when |
|---|---|
| `assets/risk_register_template.xlsx` | User wants a blank register to fill in themselves |
| `assets/risk_register_starter_kit.xlsx` | User wants something usable immediately — comes pre-filled with 15 realistic, already-scored example risks |
| `scripts/populate_risk_register.py` | Programmatically append new risk rows (see Workflow below) |

Before generating or editing the `.xlsx` file, also load `/mnt/skills/public/xlsx/SKILL.md` —
that skill defines *how* to build/edit spreadsheets correctly in this environment; this skill
defines *what* to build.

## Column schema (both files share this)

`Asset ID | Asset Name | Asset Type | Confidentiality | Integrity | Availability | Risk ID |
Risk Description | Category | Likelihood (1-5) | Impact (1-5) | Risk Score | Risk Level |
Existing Controls | Treatment Plan | Residual Likelihood (1-5) | Residual Impact (1-5) |
Residual Risk Score | Residual Risk Level | Owner | Status | Review Date`

- **Risk Score / Risk Level** = the risk BEFORE any control ("inherent risk"). Formula-driven
  (`Likelihood x Impact`, then banded Low/Medium/High/Critical). **Never overwrite this when a
  control is implemented** — it's the permanent record of what the risk originally was.
- **Residual Risk Score / Residual Risk Level** = the risk AFTER the Treatment Plan is applied.
  Also formula-driven, off separate Residual Likelihood/Impact inputs. This is where the
  improvement from a control shows up — inherent and residual stay side by side so the effect
  of a control is provable, not just claimed.
- **Confidentiality / Integrity / Availability (CIA)** rate the asset's sensitivity
  (High/Medium/Low) — explains *why* a risk to that asset matters.
- **Status** moves Open → In Progress → Mitigated/Accepted/Closed as the treatment plan is
  actually carried out (not just planned) — this is the real-world signal a control landed.

## Workflow: Add or update risks

1. Copy the appropriate template/starter kit to the working directory (don't edit in place
   unless the user gave you an existing file of theirs).
2. Turn what the user describes — even informally, e.g. "we're worried about phishing" — into
   a structured entry: Asset ID/Name/Type, CIA ratings, Risk Description, Category, Likelihood,
   Impact, Existing Controls, Treatment Plan, Owner, Status, Review Date. Ask for (or estimate
   and clearly flag as an assumption) Residual Likelihood/Impact if a treatment plan exists.
3. Write the entry as JSON and run:
   `python scripts/populate_risk_register.py <register.xlsx> <risks.json>`
   This appends a row and writes the Score/Level formulas — it never overwrites existing rows.
4. Run the xlsx skill's `recalc.py` on the file before delivering it.
5. To **update** an existing register (not create a new one), skip step 1 and pass the user's
   file directly to the script.

## Common risk categories (use if the user doesn't specify one)

Cybersecurity · Data Privacy · Third-Party / Vendor Risk · Operational · Compliance / Legal ·
Financial · Reputational · Business Continuity / Disaster Recovery

## Why this counts as "automated" (for explaining the project)

- The **template is designed once** — formulas, dropdowns (Likelihood/Impact/CIA/Status), and
  conditional formatting (color-coded risk levels, red-highlighted overdue reviews) are already
  correct. Adding a risk later is filling a row, not rebuilding a spreadsheet.
- The **script appends rows programmatically** so Risk Score/Level formulas are applied
  identically every time — no manual formula copy-paste errors.
- **Inherent vs. residual risk is tracked as two separate, always-live numbers**, so a
  control's effect is demonstrable (e.g. Critical → Medium) rather than asserted.
