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

**This workflow is for the Claude/LLM path.** If the user just wants a blank or pre-filled
register to edit themselves, hand them the template/starter kit directly and stop — don't
route a manual user through the JSON script; typing JSON isn't a value-add over typing into
Excel, and the template's dropdowns/formulas already keep manual entry correct.

1. Copy the appropriate template/starter kit to the working directory (don't edit in place
   unless the user gave you an existing file of theirs).
2. **Read the existing rows before adding anything new** (skip if the register is empty). This
   is the step that actually justifies using an LLM instead of typing a row by hand — do all
   three checks and surface findings to the user before writing the new entry:
   - **Duplicate/overlap check** — does a risk already exist for this Asset ID or Category that
     covers the same underlying cause (e.g. new "phishing" risk vs. existing "credential theft
     via social engineering")? If so, propose merging into the existing row (tightening its
     Risk Description, or adding to Existing Controls) instead of creating a near-duplicate.
   - **Control propagation check** — does the new entry's Treatment Plan (MFA, encryption at
     rest, vendor SLA, etc.) also mitigate other existing risks on the same asset that don't yet
     reflect it? If so, list those Risk IDs and ask the user whether their Residual
     Likelihood/Impact should be updated too — don't silently change other rows.
   - **Blast-radius check** — if this is a new asset that existing risks depend on (or that
     depends on an existing asset), flag whether any existing CIA ratings or Risk Descriptions
     should be revisited in light of it.

   These three checks are implemented in the script itself (text-similarity and keyword
   matching — see `scripts/populate_risk_register.py`), not just prose instructions. They are
   deliberately tuned to over-flag rather than under-flag: a flagged pair may turn out to be
   unrelated on closer reading. **The script only surfaces candidates — it never merges rows or
   edits other rows' scores on its own.** Use judgment on each flag before acting on it.
3. Turn what the user describes — even informally, e.g. "we're worried about phishing" — into
   a structured entry: Asset ID/Name/Type, CIA ratings, Risk Description, Category, Likelihood,
   Impact, Existing Controls, Treatment Plan, Owner, Status, Review Date. Ask for (or estimate
   and clearly flag as an assumption) Residual Likelihood/Impact if a treatment plan exists.
4. Write the entry as JSON (a list containing one object — see the docstring at the top of
   `populate_risk_register.py` for exact field names, e.g. `asset_id`, `description`,
   `treatment_plan`) and run the checks first (dry run, default — no `--apply` flag):
   `python scripts/populate_risk_register.py <register.xlsx> <risks.json>`
   This prints a JSON report per risk (`duplicate_check`, `control_propagation_check`,
   `blast_radius_check`) — this **is** step 2's checks, done programmatically instead of by
   eye. Review the findings and resolve/discuss with the user before writing anything. Once
   confirmed, re-run with `--apply` to actually append the row(s) and write the Score/Level
   formulas — it never overwrites existing rows.
   If step 2/4 identified other rows whose Residual Score should change, and the user confirmed
   it, update those separately and explicitly — call it out in your summary to the user rather
   than folding it silently into the same diff.
5. Run the xlsx skill's `recalc.py` on the file before delivering it.
6. To **update** an existing register (not create a new one), skip step 1 and pass the user's
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
- **Adding a risk triggers a review of the whole register, not just a new row.** Past a
  handful of entries, the actual bottleneck isn't formatting a new risk — a spreadsheet or
  human could do that — it's noticing that the new risk duplicates one already tracked, or
  that its treatment plan also lowers residual risk elsewhere, or that a new asset changes the
  blast radius of existing risks. That's a re-scan of the full sheet every single time something
  is added, which is exactly what people skip under time pressure and exactly what an LLM can
  do consistently, in one pass, at zero marginal cost per addition. This is the part of the
  value-add that scales with register size rather than shrinking to "just data entry."
