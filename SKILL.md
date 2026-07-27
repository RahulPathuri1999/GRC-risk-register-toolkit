---
name: risk-register
description: "Create, update, and maintain a GRC risk register using the organizational template (Asset Evaluation / Risk Identification / Risk Analysis / Risk Mitigation / Residual Risk Assessment schema). Use this skill whenever the user asks to create a risk register, add a risk, add a new asset, update likelihood/consequence/status/owner on a risk, or asks about risk scoring, inherent vs residual risk, or risk treatment plans. Also use when the user mentions tracking risks for an audit, a security review, or compliance work."
license: Personal project skill — not affiliated with any employer or framework body
---

# Risk Register

Helps produce and maintain a GRC risk register — a live spreadsheet, not a static document —
so risks only need to be entered once and the scoring stays correct automatically as data
changes.

| File | Use when |
|---|---|
| `assets/risk_register_template.xlsx` | User wants a blank register to fill in themselves |
| `assets/risk_register_example.xlsx` | User wants to see a realistic, filled example — 167 risks across 33 assets |
| `scripts/populate_risk_register.py` | Programmatically append new risk rows (see Workflow below) |

Before generating or editing the `.xlsx` file, also load `/mnt/skills/public/xlsx/SKILL.md` —
that skill defines *how* to build/edit spreadsheets correctly in this environment; this skill
defines *what* to build.

## Column schema

`Asset ID | Identification Date | Asset Name | Asset Owner | C | I | A | Asset Value (a) |
Risk ID | Threats | Vulnerabilities | Likelihood (b) | Consequence (c) | Risk Value (a*b*c) |
Risk Level | Risk Treatment Option | Existing Controls | Risk Mitigation Plan |
ISO 27002 Control # | Mitigation Responsibility | Mitigation Target Date |
Risk Mitigation Status | Risk Status | Residual Likelihood (d) | Residual Impact (e) |
Post-Mit Risk Value (a*d*e) | Residual Risk Level | Residual Description | Residual Control |
Risk Control Responsibility | Control Target Date | Control Status | Residual Risk Status`

- **Asset Value (a)** = Confidentiality + Integrity + Availability, each rated 1-3. Written as
  a formula (`=E{row}+F{row}+G{row}`), never hardcoded.
- **Risk Value** = Asset Value x Likelihood x Consequence — a 3-factor score, not just
  Likelihood x Impact. **Never overwrite this when a control is implemented** — it's the
  permanent record of what the risk originally was.
- **Risk Level** bands off Risk Value: Extreme (>=225) / Major (>=128) / Medium (>=37) /
  Minor (below that). Same bands apply to the residual side (Post-Mit Risk Value / Residual
  Risk Level).
- **Asset ID/Identification Date/Asset Name/Asset Owner are only filled on the FIRST row for
  a given asset** — subsequent risk rows for that asset leave those four columns blank. **C/I/A
  and Asset Value ARE repeated on every row for that asset**, per the template's own
  convention (confirmed against real filled data, not just the blank template).
- **Threats** and **Vulnerabilities** are separate columns — together they're the equivalent
  of a single "risk description" field in a simpler schema.
- **Risk Mitigation Status** and **Control Status** track the mitigation plan vs. the residual
  control separately, each with their own responsibility/target-date pair.

## Workflow: Add or update risks

**This workflow is for the Claude/LLM path.** If the user just wants a blank register to
fill in themselves, hand them `risk_register_template.xlsx` directly and stop — typing JSON
isn't a value-add over typing into Excel for a manual user.

1. Copy the appropriate file to the working directory (don't edit in place unless the user
   gave you an existing file of theirs).
2. **Read the existing rows before adding anything new** (skip if the register is empty).
   Forward-fill Asset ID/Name/Owner/CIA across blank rows first (see schema note above) so
   the checks below operate on the correct asset context per row:
   - **Duplicate/overlap check** — does a risk already exist for this Asset ID whose
     Threat + Vulnerability text describes the same underlying cause? If so, propose merging
     into the existing row instead of creating a near-duplicate.
   - **Control propagation check** — does the new entry's Mitigation Plan also mitigate other
     existing risks on the same asset that don't yet reflect it? If so, list those Risk IDs
     and ask the user whether their residual fields should be updated too.
   - **Blast-radius check** — if this is a new asset connected to an existing one, flag
     whether any existing CIA ratings or Threats/Vulnerabilities should be revisited.

   These checks are implemented in the script itself (text-similarity and keyword matching —
   see `scripts/populate_risk_register.py`), not just prose instructions. They're tuned for
   checking *one new risk* against the register — running them as a full pairwise audit
   across an already-large register needs a stricter similarity threshold (~0.3 rather than
   the default same-asset bar) or it gets noisy; this was confirmed by testing against a real
   167-risk register. **The script only surfaces candidates — it never merges rows or edits
   other rows' scores on its own.** Use judgment on each flag before acting on it.
3. Turn what the user describes — even informally, e.g. "we're worried about phishing" —
   into a structured entry: Asset ID/Name/Owner, CIA ratings, Threat, Vulnerability,
   Likelihood, Consequence, Existing Controls, Mitigation Plan, Owner. Ask for (or estimate
   and clearly flag as an assumption) Residual Likelihood/Impact if a mitigation plan exists.
4. Write the entry as JSON (a list containing one object — see the docstring at the top of
   `populate_risk_register.py` for exact field names) and run the checks first (dry run,
   default — no `--apply` flag):
   `python scripts/populate_risk_register.py <register.xlsx> <risks.json>`
   This prints a JSON report per risk (`duplicate_check`, `control_propagation_check`,
   `blast_radius_check`). Review the findings and resolve/discuss with the user before
   writing anything. Once confirmed, re-run with `--apply` to actually append the row(s) and
   write the Risk Value/Level formulas — it never overwrites existing rows.
   If step 2/4 identified other rows whose residual fields should change, and the user
   confirmed it, update those separately and explicitly — call it out in your summary rather
   than folding it silently into the same diff.
5. Run the xlsx skill's `recalc.py` on the file before delivering it.
6. To **update** an existing register (not create a new one), skip step 1 and pass the user's
   file directly to the script.
7. **If the user's existing file has Likelihood/Consequence filled in but Risk Value/Level
   left blank** (a real, common gap — 62% of rows in the example register had this), offer to
   backfill those as live formulas rather than leaving them broken. Also check for duplicate
   Risk IDs across the whole file, not just within one asset — IDs must be globally unique.

## Why this counts as "automated" (for explaining the project)

- **Risk Value / Risk Level / Post-Mit Risk Value / Residual Risk Level are Excel formulas**,
  not typed-in numbers — the sheet recalculates itself if inputs change. The template as
  originally provided had none of these as formulas; adding them closes a real gap (found
  103 rows in a real 167-risk register missing calculated values despite having the inputs).
- **The script appends rows programmatically**, correctly handling the template's
  first-row-per-asset convention, so Risk Value/Level formulas are applied identically every
  time — no manual formula copy-paste errors, and no reused Risk IDs (a real bug found and
  fixed in testing).
- **Inherent vs. residual risk is tracked as two separate, always-live scores**, so a
  control's effect is demonstrable rather than asserted.
- **Adding a risk triggers a review of the whole register, not just a new row.** Past a
  handful of entries, the actual bottleneck isn't formatting a new risk — it's noticing that
  it duplicates one already tracked, or that its mitigation plan also lowers risk elsewhere,
  or that a new asset changes the blast radius of existing risks. That's a re-scan of the
  full sheet every time something is added, which is exactly what people skip under time
  pressure and exactly what an LLM can do consistently, in one pass, at zero marginal cost
  per addition.
