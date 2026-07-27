"""
Append new risks to a risk register built from risk_register_template.xlsx or
risk_register_starter_kit.xlsx (both share the same column schema), with pre-flight
checks that catch what a person re-scanning the whole sheet by eye would look for:

  1. duplicate_check          - existing risks on the same asset with similar wording
  2. control_propagation_check- existing risks whose residual score might also improve
  3. blast_radius_check       - flags when a new asset connects to ones already tracked

Usage:
    python populate_risk_register.py <path_to_register.xlsx> <path_to_risks.json> [--apply]

Without --apply (default): prints a JSON report of findings for each risk and does NOT
write anything. Review the report, then re-run with --apply to actually append the rows.

risks.json format (a list of objects):
[
  {
    "asset_id": "AST-016",
    "asset_name": "Development Environment",
    "asset_type": "Server",
    "confidentiality": "Medium",
    "integrity": "Medium",
    "availability": "Low",
    "risk_id": "R-016",
    "description": "Production credentials hardcoded in a shared dev repository",
    "category": "Cybersecurity",
    "likelihood": 3,
    "impact": 4,
    "existing_controls": "Code review process (not consistently enforced)",
    "treatment_plan": "Move secrets to a vault (e.g. HashiCorp Vault); add pre-commit secret scanning",
    "residual_likelihood": 1,
    "residual_impact": 4,
    "owner": "Engineering",
    "status": "Open",
    "review_date": "2026-10-01"
  }
]

All fields are optional except a meaningful "description". Confidentiality/Integrity/
Availability accept "High"/"Medium"/"Low". Residual fields are optional -- leave them out if a
treatment plan isn't decided yet. Risk Score, Risk Level, Residual Risk Score, and Residual
Risk Level are all written as formulas (not computed in Python), so the sheet keeps
recalculating itself if any Likelihood/Impact value is edited later.
"""
import sys
import json
import re
import argparse
from difflib import SequenceMatcher
import openpyxl

# column letters, 1-indexed to match the sheet
COLS = {
    "asset_id": 1, "asset_name": 2, "asset_type": 3, "confidentiality": 4, "integrity": 5,
    "availability": 6, "risk_id": 7, "description": 8, "category": 9, "likelihood": 10,
    "impact": 11, "risk_score": 12, "risk_level": 13, "existing_controls": 14,
    "treatment_plan": 15, "residual_likelihood": 16, "residual_impact": 17,
    "residual_risk_score": 18, "residual_risk_level": 19, "owner": 20, "status": 21,
    "review_date": 22,
}

STOPWORDS = {
    "the", "a", "an", "to", "of", "in", "on", "could", "would", "and", "or", "due",
    "via", "for", "is", "are", "be", "by", "with", "into", "leading", "that",
}

CONTROL_KEYWORDS = {
    "mfa": {"password", "credential", "login", "authentication", "access"},
    "multi-factor": {"password", "credential", "login", "authentication", "access"},
    "encryption": {"data", "database", "storage", "laptop", "device"},
    "training": {"phishing", "social", "employee", "staff", "human"},
    "awareness": {"phishing", "social", "employee", "staff", "human"},
    "backup": {"availability", "outage", "disaster", "downtime"},
    "sla": {"vendor", "third-party", "breach", "notification"},
}


def tokenize(text):
    words = re.findall(r"[a-z]+", (text or "").lower())
    return {w for w in words if w not in STOPWORDS and len(w) > 2}


def text_similarity(a, b):
    ta, tb = tokenize(a), tokenize(b)
    jaccard = len(ta & tb) / len(ta | tb) if (ta or tb) else 0
    seq = SequenceMatcher(None, a or "", b or "").ratio()
    return round(0.7 * jaccard + 0.3 * seq, 2)


def first_empty_row(ws):
    """Find the first row with no Risk ID -- the template/starter kit are pre-formatted
    hundreds of rows down, so ws.max_row is NOT a safe place to append."""
    row = 2
    while ws.cell(row=row, column=COLS["risk_id"]).value not in (None, ""):
        row += 1
    return row


def load_existing_rows(ws, ws_vals, up_to_row):
    rows = []
    for r in range(2, up_to_row):
        if ws.cell(row=r, column=COLS["risk_id"]).value in (None, ""):
            continue
        row = {k: ws.cell(row=r, column=c).value for k, c in COLS.items()}
        row["risk_level"] = ws_vals.cell(row=r, column=COLS["risk_level"]).value
        row["residual_risk_level"] = ws_vals.cell(row=r, column=COLS["residual_risk_level"]).value
        row["_row"] = r
        rows.append(row)
    return rows


def duplicate_check(new_risk, existing_rows, threshold=0.35, same_category_threshold=0.12):
    """Same-asset risks are flagged at a lower bar if they also share a Category --
    a shared category plus even a single strong keyword overlap (e.g. "phishing") is a
    much stronger duplicate signal than raw word overlap alone, and plain Jaccard similarity
    badly under-scores paraphrases (different wording, same underlying risk)."""
    flags = []
    for row in existing_rows:
        if row["asset_id"] != new_risk.get("asset_id"):
            continue
        sim = text_similarity(new_risk.get("description"), row["description"])
        same_category = row.get("category") == new_risk.get("category")
        bar = same_category_threshold if same_category else threshold
        if sim >= bar:
            flags.append({
                "risk_id": row["risk_id"], "description": row["description"], "similarity": sim,
                "same_category": same_category,
            })
    flags.sort(key=lambda f: -f["similarity"])
    return flags


def control_propagation_check(new_risk, existing_rows):
    plan_text = (new_risk.get("treatment_plan") or "").lower()
    matched = [kw for kw in CONTROL_KEYWORDS if kw in plan_text]
    if not matched:
        return []
    related_terms = set().union(*(CONTROL_KEYWORDS[kw] for kw in matched))
    flags = []
    for row in existing_rows:
        if row["asset_id"] != new_risk.get("asset_id") or row["risk_id"] == new_risk.get("risk_id"):
            continue
        if tokenize(row["description"]) & related_terms:
            flags.append({
                "risk_id": row["risk_id"], "description": row["description"],
                "current_residual_level": row["residual_risk_level"],
                "reason": f"Treatment plan mentions {matched}, which plausibly also mitigates this risk's cause",
            })
    return flags


GENERIC_ASSET_WORDS = {
    "company", "corporate", "employee", "employees", "system", "systems", "data",
    "service", "services", "management", "platform", "process", "application", "applications",
}


def blast_radius_check(new_risk, existing_rows):
    is_new_asset = not any(r["asset_id"] == new_risk.get("asset_id") for r in existing_rows)
    if not is_new_asset:
        return None
    new_tokens = tokenize(new_risk.get("asset_name", "")) - GENERIC_ASSET_WORDS
    related, seen = [], set()
    for row in existing_rows:
        if row["asset_id"] in seen:
            continue
        candidate_tokens = (tokenize(row.get("asset_name", "")) | tokenize(row.get("description", ""))) - GENERIC_ASSET_WORDS
        if new_tokens & candidate_tokens:
            related.append(row["asset_id"])
            seen.add(row["asset_id"])
    return {"new_asset": True, "possibly_related_existing_assets": related}


def run_checks(new_risk, existing_rows):
    return {
        "risk_id": new_risk.get("risk_id"),
        "duplicate_check": duplicate_check(new_risk, existing_rows),
        "control_propagation_check": control_propagation_check(new_risk, existing_rows),
        "blast_radius_check": blast_radius_check(new_risk, existing_rows),
    }


def write_row(ws, row, risk, fallback_id):
    ws.cell(row=row, column=COLS["asset_id"], value=risk.get("asset_id", ""))
    ws.cell(row=row, column=COLS["asset_name"], value=risk.get("asset_name", ""))
    ws.cell(row=row, column=COLS["asset_type"], value=risk.get("asset_type", ""))
    ws.cell(row=row, column=COLS["confidentiality"], value=risk.get("confidentiality", ""))
    ws.cell(row=row, column=COLS["integrity"], value=risk.get("integrity", ""))
    ws.cell(row=row, column=COLS["availability"], value=risk.get("availability", ""))
    ws.cell(row=row, column=COLS["risk_id"], value=risk.get("risk_id", fallback_id))
    ws.cell(row=row, column=COLS["description"], value=risk.get("description", ""))
    ws.cell(row=row, column=COLS["category"], value=risk.get("category", ""))
    ws.cell(row=row, column=COLS["likelihood"], value=risk.get("likelihood"))
    ws.cell(row=row, column=COLS["impact"], value=risk.get("impact"))
    ws.cell(row=row, column=COLS["risk_score"], value=f"=J{row}*K{row}")
    ws.cell(row=row, column=COLS["risk_level"],
            value=f'=IF(L{row}>=15,"Critical",IF(L{row}>=9,"High",IF(L{row}>=4,"Medium","Low")))')
    ws.cell(row=row, column=COLS["existing_controls"], value=risk.get("existing_controls", ""))
    ws.cell(row=row, column=COLS["treatment_plan"], value=risk.get("treatment_plan", ""))
    ws.cell(row=row, column=COLS["residual_likelihood"], value=risk.get("residual_likelihood"))
    ws.cell(row=row, column=COLS["residual_impact"], value=risk.get("residual_impact"))
    ws.cell(row=row, column=COLS["residual_risk_score"], value=f"=P{row}*Q{row}")
    ws.cell(row=row, column=COLS["residual_risk_level"],
            value=f'=IF(R{row}>=15,"Critical",IF(R{row}>=9,"High",IF(R{row}>=4,"Medium","Low")))')
    ws.cell(row=row, column=COLS["owner"], value=risk.get("owner", ""))
    ws.cell(row=row, column=COLS["status"], value=risk.get("status", "Open"))
    ws.cell(row=row, column=COLS["review_date"], value=risk.get("review_date", ""))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx_path")
    ap.add_argument("json_path")
    ap.add_argument("--apply", action="store_true", help="Actually write the rows (default: dry run / report only)")
    args = ap.parse_args()

    with open(args.json_path) as f:
        risks = json.load(f)
    if isinstance(risks, dict):
        risks = [risks]

    wb = openpyxl.load_workbook(args.xlsx_path)
    ws = wb["Risk Register"]
    wb_vals = openpyxl.load_workbook(args.xlsx_path, data_only=True)
    ws_vals = wb_vals["Risk Register"]

    start_row = first_empty_row(ws)
    existing_rows = load_existing_rows(ws, ws_vals, start_row)

    reports = []
    row = start_row
    for i, risk in enumerate(risks):
        report = run_checks(risk, existing_rows)
        reports.append(report)
        fallback_id = f"R-{row - 1:03d}"
        if args.apply:
            write_row(ws, row, risk, fallback_id)
        # Always add this risk to existing_rows (even in dry-run) so later risks in the SAME
        # batch are checked against it too -- otherwise a dry-run preview wrongly misses
        # duplicates/overlaps between risks submitted together in one file.
        existing_rows.append({
            **{k: risk.get(k) for k in COLS},
            "risk_id": risk.get("risk_id", fallback_id),
            "residual_risk_level": None,
            "_row": row,
        })
        if args.apply:
            row += 1

    print(json.dumps(reports, indent=2))

    if args.apply:
        wb.save(args.xlsx_path)
        print(f"\nAppended {len(risks)} risk(s) to {args.xlsx_path}. Remember to run recalc.py.", file=sys.stderr)


if __name__ == "__main__":
    main()
