# Risk Register Toolkit

A GRC risk register built around a real organizational template (Asset Evaluation → Risk
Identification → Risk Analysis → Risk Mitigation → Residual Risk Assessment), with a script
that adds new risks safely and catches problems a person re-scanning the sheet by hand would
otherwise miss.

## Files

```
├── SKILL.md                              # Instructions Claude reads to use this
├── README.md                             # This file
├── assets/
│   ├── risk_register_template.xlsx       # Blank, ready to fill in
│   └── risk_register_example.xlsx        # A cleaned, real 167-risk / 33-asset register
└── scripts/
    └── populate_risk_register.py         # Appends new risk rows with pre-flight checks
```

## Why this exists

Filling in a risk register is easy for the first few entries. It gets harder to keep correct
as it grows, because three things start happening that a spreadsheet won't catch on its own:

1. **Risk Value/Level get typed by hand and forgotten.** Running this on a real 167-risk
   register turned up **103 rows (62%) where Likelihood and Consequence were filled in but
   the Risk Value/Level were never calculated** — someone filled in the inputs and moved on
   before finishing the formula. The template as originally built had *no formulas at all*
   for these fields; this project adds them as real Excel formulas
   (`Asset Value x Likelihood x Consequence`, banded into Minor/Medium/Major/Extreme) so this
   can't happen again for new rows.
2. **IDs collide.** The same real register had a Risk ID (`R-001`) reused across two
   unrelated assets — invisible in a 167-row sheet, obvious once you know to check for it.
3. **Duplicate or overlapping risks creep in.** Two risks phrased differently can describe
   the same underlying problem. Re-reading a large sheet to catch that by eye doesn't scale;
   this project's pre-flight checks do that scan automatically every time a risk is added.

## Quick use (in Claude)

1. Upload this as a zip: claude.ai → Settings → Customize → Skills → Upload skill
2. Make sure the toggle is ON
3. In a new chat: *"Add a risk to my risk register: [describe it]"*

## Quick use (standalone, no Claude)

```bash
pip install openpyxl
python scripts/populate_risk_register.py assets/risk_register_template.xlsx your_risks.json
```
By default this is a **dry run** — it prints a JSON report of pre-flight checks and writes
nothing. Review it, then re-run with `--apply` to actually write the rows:
```bash
python scripts/populate_risk_register.py assets/risk_register_template.xlsx your_risks.json --apply
```
See the docstring at the top of `populate_risk_register.py` for the full JSON field list.

## Pre-flight checks

Before writing anything, the script scans the existing register (forward-filling Asset
ID/Name/Owner/CIA, since the template only fills those on the first row of each asset group)
and flags:

- **Duplicate check** — an existing risk on the same asset with similar Threat/Vulnerability
  wording, even if phrased differently.
- **Control propagation check** — an existing risk on the same asset whose Mitigation Plan
  keywords (MFA, encryption, patching, backup, access controls, etc.) suggest the new risk's
  fix would also help it.
- **Blast radius check** — if the new risk introduces a brand-new asset, flags existing
  assets it's plausibly connected to.

These are candidates for review, not verdicts — a flagged pair can turn out to be unrelated
on a closer read, and the script never merges rows or edits other rows' scores on its own.
The same-asset duplicate-check threshold is deliberately loose for the single-new-risk-add
workflow this script is built for; running it as a full pairwise audit across an entire
existing register (as opposed to checking one new addition) needs a stricter threshold to
avoid noise — worth knowing if you extend this into a standalone audit mode.

## What makes this "automated"

- **Risk Value / Risk Level / Post-Mit Risk Value / Residual Risk Level are Excel formulas**,
  not typed-in numbers — the sheet recalculates itself if any input changes, and can't end up
  in the half-finished state 103 rows of the real example register were found in.
- **Asset ID/Name/Owner/CIA only need to be entered once per asset** — the script correctly
  reads (and writes) the template's convention of leaving those blank on subsequent risk rows
  for the same asset, while still repeating C/I/A/Asset Value on every row.
- **Adding a risk triggers a review of the whole register, not just a new row** — the
  pre-flight checks re-scan for duplicates, control overlap, and asset relationships every
  time, which is exactly the kind of check that gets skipped under time pressure when done
  by hand.
