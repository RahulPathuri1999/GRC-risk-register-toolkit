"""
Append new risks to a risk register built from risk_register_template.xlsx or
risk_register_starter_kit.xlsx (both share the same column schema).

Usage:
    python populate_risk_register.py <path_to_register.xlsx> <path_to_risks.json>

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
import openpyxl


def main():
    if len(sys.argv) != 3:
        print("Usage: python populate_risk_register.py <register.xlsx> <risks.json>")
        sys.exit(1)

    xlsx_path, json_path = sys.argv[1], sys.argv[2]

    with open(json_path) as f:
        risks = json.load(f)

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb["Risk Register"]

    # Find first empty row after existing data (search on Risk ID, column G)
    row = 2
    while ws.cell(row=row, column=7).value not in (None, ""):
        row += 1

    for risk in risks:
        ws.cell(row=row, column=1, value=risk.get("asset_id", ""))
        ws.cell(row=row, column=2, value=risk.get("asset_name", ""))
        ws.cell(row=row, column=3, value=risk.get("asset_type", ""))
        ws.cell(row=row, column=4, value=risk.get("confidentiality", ""))
        ws.cell(row=row, column=5, value=risk.get("integrity", ""))
        ws.cell(row=row, column=6, value=risk.get("availability", ""))
        ws.cell(row=row, column=7, value=risk.get("risk_id", f"R-{row-1:03d}"))
        ws.cell(row=row, column=8, value=risk.get("description", ""))
        ws.cell(row=row, column=9, value=risk.get("category", ""))
        ws.cell(row=row, column=10, value=risk.get("likelihood"))
        ws.cell(row=row, column=11, value=risk.get("impact"))
        ws.cell(row=row, column=12, value=f"=J{row}*K{row}")
        ws.cell(row=row, column=13,
                value=f'=IF(L{row}>=15,"Critical",IF(L{row}>=9,"High",IF(L{row}>=4,"Medium","Low")))')
        ws.cell(row=row, column=14, value=risk.get("existing_controls", ""))
        ws.cell(row=row, column=15, value=risk.get("treatment_plan", ""))
        ws.cell(row=row, column=16, value=risk.get("residual_likelihood"))
        ws.cell(row=row, column=17, value=risk.get("residual_impact"))
        ws.cell(row=row, column=18, value=f"=P{row}*Q{row}")
        ws.cell(row=row, column=19,
                value=f'=IF(R{row}>=15,"Critical",IF(R{row}>=9,"High",IF(R{row}>=4,"Medium","Low")))')
        ws.cell(row=row, column=20, value=risk.get("owner", ""))
        ws.cell(row=row, column=21, value=risk.get("status", "Open"))
        ws.cell(row=row, column=22, value=risk.get("review_date", ""))
        row += 1

    wb.save(xlsx_path)
    print(f"Added {len(risks)} risk(s) to {xlsx_path}. Remember to run recalc.py.")


if __name__ == "__main__":
    main()
