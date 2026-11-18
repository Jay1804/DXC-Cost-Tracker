"""DXC Cost Tracker builder.

Runs the "Cost Approval" and "Client Name Disclosure" queries against the live DB,
appends the results and writes a formatted Excel tracker.

Usage:  python build_tracker.py
Reads DB_HOST / DB_NAME / DB_USER / DB_PASSWORD from .env next to this script.
"""
from datetime import datetime
from pathlib import Path

import pandas as pd
import pymysql
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).parent
# DXC client accounts that carry the DXC flex fields (Candidate ID, Package Name, ...)
DXC_CLIENT_IDS = (158686, 398650)

FINAL_COLUMNS = [
    "Case_Check_id", "case_ars_no", "Company_name", "Candidate_name", "Process_name",
    "Candidate ID", "VS City", "Case initiator Email Id", "location", "case_received_date",
    "check_name", "VS Name", "insuff_remarks", "Cost", "Package Name", "Pre Hire or Post Hire",
]

FLEX = ",".join(f"ecf.CASE_FLEX_FIELD{i}" for i in range(1, 31))
IDS = ",".join(str(i) for i in DXC_CLIENT_IDS)

COMMON_JOINS = """
FROM ec_case_master ecm
LEFT JOIN ec_case_checks ecc ON ecm.case_id = ecc.case_id
LEFT JOIN ec_master_company emc ON emc.company_id = ecm.client_id
LEFT JOIN ec_client_process ecp ON ecm.process_id=ecp.process_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id=ecm.candidate_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id=emcl.office_id
LEFT JOIN ec_case_check_verification_source eccvs ON ecc.case_check_id=eccvs.case_check_id
LEFT JOIN ec_master_educational_institute emei ON eccvs.org_id=emei.institute_id
LEFT JOIN ec_master_company emc1 ON eccvs.org_id=emc1.company_id
LEFT JOIN ec_master_city emc4 ON emc1.city_id=emc4.city_id
LEFT JOIN ec_master_city emc6 ON emei.city_id=emc6.city_id
LEFT JOIN ec_case_fields ecf ON ecm.case_id=ecf.case_id
LEFT JOIN ec_checks ec ON ecc.check_id=ec.check_id
WHERE ecc.check_status=3 AND case_status NOT IN (6,8,14)
AND ecm.client_id IN ({ids})
AND first_name NOT LIKE '%%dummy%%' AND first_name NOT LIKE '%%test%%'
"""

SELECT_COMMON = f"""
SELECT ecm.client_id AS client_id,
    ecc.Case_Check_id, case_ars_no, emc.Company_name,
    CONCAT(first_name,' ',IFNULL(Middle_name,''),' ',last_name) AS Candidate_name,
    Process_name, office_name AS location, received_date AS case_received_date,
    ecc.check_name AS check_name,
    ecc.insuff_remarks AS insuff_remarks,
    checkpoint_live.fn_vs_name(ecc.case_check_id) AS `VS Name`,
    IF(ecc.family_id=5, emc4.city_name, IF(ecc.family_id=4, emc6.city_name, '')) AS `VS City`,
    {FLEX}
"""

COST_APPROVAL_SQL = (
    SELECT_COMMON
    + """,
    (SELECT cost FROM ec_case_insuff_details ecid
        WHERE ecc.case_check_id=ecid.case_check_id AND insuff_raised_date IS NOT NULL
        ORDER BY insuff_id DESC LIMIT 1) AS Cost
"""
    + COMMON_JOINS.format(ids=IDS)
    + " AND (ecc.insuff_remarks LIKE '%%cost%%' OR ecc.insuff_remarks LIKE '%%approval%%')"
)

DISCLOSURE_SQL = (
    SELECT_COMMON
    + ", NULL AS Cost\n"
    + COMMON_JOINS.format(ids=IDS)
    + " AND ecc.insuff_remarks LIKE '%%disclosure%%'"
)

FIELD_MAP_SQL = """SELECT client_id, field_id, field_name FROM ec_client_case_fields
WHERE client_id IN ({ids})""".format(ids=IDS)


def load_env():
    env = {}
    for line in (HERE / ".env").read_text().splitlines():
        if "=" in line and not line.lstrip().startswith("#"):
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip("\"'")
    return env


def fetch(cur, sql):
    cur.execute(sql)
    cols = [d[0] for d in cur.description]
    return pd.DataFrame(cur.fetchall(), columns=cols)


def resolve_flex(df, field_map):
    """Pull client-specific flex fields (by field name) into named columns."""
    wanted = ["Candidate ID", "Case initiator Email Id", "Package Name", "Pre Hire or Post Hire"]
    for w in wanted:
        df[w] = None
    for cid, grp in field_map.groupby("client_id"):
        names = {r.field_name.strip().lower(): int(r.field_id) for r in grp.itertuples()}
        mask = df["client_id"] == cid
        for w in wanted:
            fid = names.get(w.lower())
            if fid:
                df.loc[mask, w] = df.loc[mask, f"CASE_FLEX_FIELD{fid}"]
    return df


def write_excel(df, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Cost Tracker"
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill = PatternFill("solid", start_color="D3D3D3", end_color="D3D3D3")

    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append([None if pd.isna(v) else v for v in row])

    for c in ws[1]:
        c.font = Font(bold=True, name="Calibri", size=11)
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for c in row:
            c.border = border
            if c.row > 1:
                c.font = Font(name="Calibri", size=10)
                c.alignment = Alignment(vertical="top")
            if hasattr(c.value, "year"):
                c.number_format = "dd-mmm-yyyy"

    for i, col in enumerate(df.columns, 1):
        longest = max([len(str(col))] + [len(str(v)) for v in df[col].head(500) if pd.notna(v)])
        ws.column_dimensions[get_column_letter(i)].width = min(max(12, longest + 2), 50)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def build_df():
    env = load_env()
    conn = pymysql.connect(host=env["DB_HOST"], user=env["DB_USER"], password=env["DB_PASSWORD"],
                           database=env["DB_NAME"], connect_timeout=30, read_timeout=900)
    cur = conn.cursor()
    cost = fetch(cur, COST_APPROVAL_SQL)
    disc = fetch(cur, DISCLOSURE_SQL)
    fmap = fetch(cur, FIELD_MAP_SQL)
    conn.close()
    df = pd.concat([cost, disc], ignore_index=True)
    df = resolve_flex(df, fmap)
    df = df[FINAL_COLUMNS].drop_duplicates()
    return df.sort_values(["case_received_date", "Case_Check_id"], ascending=[False, True])


def main():
    df = build_df()
    out = HERE / f"DXC_Cost_Tracker_{datetime.now():%Y%m%d_%H%M}.xlsx"
    write_excel(df, out)
    print(f"Tracker rows: {len(df)}\nSaved: {out}")


if __name__ == "__main__":
    main()
