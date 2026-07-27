# Learnings & Design Decisions

Notes on why this project exists, what it found, and what I'd change next — written as
documentation, not as a pitch.

## The problem

A risk register is easy to keep correct for the first 10-15 entries. Past that, three things
start going wrong that a spreadsheet doesn't catch on its own:

1. **Formulas get typed by hand and forgotten.** Someone fills in Likelihood and Consequence,
   gets interrupted, and the Risk Value cell stays blank. Across a year of entries, that adds
   up to a register where a large share of rows don't actually have a usable score.
2. **IDs collide.** Nothing stops the same Risk ID from being reused for a different asset in
   a long sheet. It doesn't break anything visibly — it just quietly makes the register
   unreliable as a reference.
3. **Duplicate risks creep in.** Two entries can describe the same underlying risk in
   different words. Nobody catches it because nobody re-reads the whole sheet every time a
   row is added.

## What this project does about it

- **The template has real formulas**, not numbers someone has to calculate and type in. Risk
  Value, Risk Level, Post-Mitigation Risk Value, and Residual Risk Level are all Excel
  formulas (`Asset Value × Likelihood × Consequence`, banded into Minor/Medium/Major/Extreme).
  The template as originally provided had none of these as formulas.
- **`populate_risk_register.py` appends new risks with three pre-flight checks**, run before
  anything is written:
  - **Duplicate check** — flags an existing risk on the same asset with similar
    Threat/Vulnerability wording.
  - **Control propagation check** — if the new risk's mitigation plan would plausibly also
    fix an existing risk on the same asset that hasn't been updated to reflect it, that row
    gets flagged.
  - **Blast-radius check** — adding a new asset flags which existing assets it's plausibly
    connected to.
- **Dry-run by default.** The script prints its findings as a JSON report first; nothing is
  written until it's re-run with `--apply`.

## What testing against a real register found

I ran this against a real, already-filled 167-risk register across 33 assets before writing
any of this up as "done." It found:

- **103 rows (62%) had Likelihood/Consequence filled in but no calculated Risk Value or Risk
  Level.** The inputs were present; the output was silently missing.
- **One Risk ID (`R-001`) was reused across two unrelated assets.**
- Two genuinely borderline near-duplicate risks worth a second look — the tool flagged them,
  it didn't merge them automatically.

`assets/risk_register_example.xlsx` is that register, cleaned: the 103 formulas backfilled,
the duplicate ID renumbered, and names anonymized before publishing.

## Where the value actually is (and where it isn't)

For someone filling in one or two risks by hand, this doesn't save much time — typing a row
into Excel is about as fast as writing JSON, and the value there is really just the template's
formulas keeping the score correct.

The value scales with the size of the register. Past a certain point, the bottleneck stops
being data entry and becomes review: noticing a new risk duplicates one already tracked, or
that its fix also lowers residual risk elsewhere, or that a new asset changes the picture for
existing ones. Re-scanning a 100+ row sheet by eye every time something is added is exactly
the kind of check that gets skipped under time pressure. The script does that scan
automatically, every time, at no added cost per addition.

## Known limitations

- **The duplicate check is keyword/text-similarity based, not semantic.** It can miss
  paraphrases (different wording, same underlying risk) and can false-positive on incidental
  word overlap (e.g. flagging two unrelated assets because both descriptions happen to share
  a common word). The threshold is tuned to over-flag rather than under-flag, and the script
  never auto-merges — it only ever surfaces a candidate for review.
- **The threshold tuned for "check one new risk against the register" is too loose for "audit
  the whole register at once."** Running a full pairwise scan across the 167-risk register at
  the same threshold produced far too many weak matches. A bulk-audit mode would need its own,
  stricter threshold — not built yet.
- **This is built for one specific organizational schema**, not a generic tool. The column
  layout, the first-row-per-asset convention, and the 3-factor scoring
  (`Asset Value × Likelihood × Consequence` rather than just `Likelihood × Impact`) all match
  how this particular template works.

## Next steps

- A standalone audit mode (higher similarity threshold, run once across an existing register
  rather than per-addition).
- Semantic similarity instead of keyword/Jaccard matching for the duplicate check, to catch
  paraphrases the current version misses.
- Validation on ingest that rejects a batch introducing a duplicate Risk ID before it reaches
  the sheet, instead of relying on a manual audit afterward.
