"""
Script to generate the comprehensive BSNL 5-Invoice Reconciliation & Outage Audit Workbook.
Covers:
- Invoice 1: 607239111046515 (100 Sites, PDF Pages 1-5)
- Invoice 2: 607239111046516 (85 Sites, PDF Pages 6-9)
- Invoice 3: 607239111046519 (99 Sites, PDF Pages 10-14)
- Invoice 4: 607239111046520 (100 Sites, PDF Pages 15-19)
- Invoice 5: 607239111046533 (100 Sites, PDF Pages 20-24)
- Outage Sign-off matching (PDF Pages 25-80)
"""

import os
import xlrd
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# MSA SLA Constants
ALLOWED_OUTAGE_MINS = 40.32  # 28 days * 24 hrs * 60 mins * (1 - 0.999)
PENALTY_RATE_PER_HR = 500.0
PENALTY_CAP = 10000.0

# 1. Load data from the master source sheet
source_xls = "media_1794423587827_bef8daec.xls"
book = xlrd.open_workbook(source_xls)
sheet = book.sheet_by_name("FEB-26")

# Read rows
rows_data = []
for r in range(2, sheet.nrows):
    row = sheet.row_values(r)
    inv = row[14]
    if not inv:
        continue
    try:
        inv_str = str(int(float(inv)))
    except:
        inv_str = str(inv).strip()

    gid = row[4]
    try:
        gid_str = str(int(float(gid)))
    except:
        gid_str = str(gid).strip()

    site_name = str(row[3]).strip()
    vendor_id = str(row[5]).strip()
    bsnl_id = str(row[6]).strip()
    ssa = str(row[8]).strip()
    eqpt_type = str(row[11]).strip()
    bts_eqpt = str(row[13]).strip()

    # Outage mins
    outage_min = row[15]
    try:
        outage_min = float(outage_min) if outage_min != "" else 0.0
    except:
        outage_min = 0.0

    # Tenancy
    tenancy = row[17]
    try:
        tenancy = int(float(tenancy)) if tenancy != "" else 2
    except:
        tenancy = 2

    claimed_tenancy = row[18]
    try:
        claimed_tenancy = int(float(claimed_tenancy)) if claimed_tenancy != "" else tenancy
    except:
        claimed_tenancy = tenancy

    ssa_type = str(row[20]).strip()
    claimed_city = str(row[21]).strip()
    dg_cap = row[22]
    try:
        dg_cap = str(int(float(dg_cap))) if dg_cap != "" else "Non DG"
    except:
        dg_cap = str(dg_cap).strip()

    bts_rental_type = str(row[23]).strip()

    # Energy parameters
    dg_run = row[25]
    try:
        dg_run = float(dg_run) if dg_run != "" else 0.0
    except:
        dg_run = 0.0

    dg_cph = row[27]
    try:
        dg_cph = float(dg_cph) if dg_cph != "" else 0.0
    except:
        dg_cph = 0.0

    diesel_rate = row[28]
    try:
        diesel_rate = float(diesel_rate) if diesel_rate != "" else 91.88
    except:
        diesel_rate = 91.88

    diesel_chg_day = row[29]
    try:
        diesel_chg_day = float(diesel_chg_day) if diesel_chg_day != "" else 0.0
    except:
        diesel_chg_day = 0.0

    eb_hrs = row[30]
    try:
        eb_hrs = float(eb_hrs) if eb_hrs != "" else 24.0
    except:
        eb_hrs = 24.0

    eb_rate = row[31]
    try:
        eb_rate = float(eb_rate) if eb_rate != "" else 8.40
    except:
        eb_rate = 8.40

    eb_cph = row[32]
    try:
        eb_cph = float(eb_cph) if eb_cph != "" else 0.0
    except:
        eb_cph = 0.0

    eb_chg_day = row[33]
    try:
        eb_chg_day = float(eb_chg_day) if eb_chg_day != "" else 0.0
    except:
        eb_chg_day = 0.0

    total_chg_day = row[34]
    try:
        total_chg_day = float(total_chg_day) if total_chg_day != "" else 0.0
    except:
        total_chg_day = 0.0

    total_chg_month = row[35]
    try:
        total_chg_month = float(total_chg_month) if total_chg_month != "" else 0.0
    except:
        total_chg_month = 0.0

    claimed_amount = row[36]
    try:
        claimed_amount = float(claimed_amount) if claimed_amount != "" else 0.0
    except:
        claimed_amount = 0.0

    passed_base = row[37]
    try:
        passed_base = float(passed_base) if passed_base != "" else claimed_amount
    except:
        passed_base = claimed_amount

    # Deduction
    deduction = row[41]
    try:
        deduction = float(deduction) if deduction != "" else 0.0
    except:
        deduction = 0.0

    remark = str(row[40]).strip()
    if not remark and len(row) > 42:
        remark = str(row[42]).strip()

    # Calculate SLA Outage Penalty
    excess_mins = 0.0
    outage_penalty = 0.0
    outage_note = "Uptime 100% (No outage)"

    if outage_min > 0:
        if outage_min <= ALLOWED_OUTAGE_MINS:
            outage_note = f"Outage {outage_min}m within allowed {ALLOWED_OUTAGE_MINS}m"
        else:
            excess_mins = outage_min - ALLOWED_OUTAGE_MINS
            raw_pen = (excess_mins / 60.0) * PENALTY_RATE_PER_HR
            pen = max(500.0, raw_pen)
            outage_penalty = min(pen, PENALTY_CAP)
            is_cap = pen > PENALTY_CAP
            outage_note = f"Downtime {outage_min:.0f}m (Excess {excess_mins:.1f}m / {excess_mins/60.0:.2f}h). Penalty ₹{outage_penalty:,.2f}" + (" [Capped at ₹10k max]" if is_cap else "")

    # Net Approved = passed_base - outage_penalty
    net_approved = max(0.0, passed_base - outage_penalty)
    total_ded = claimed_amount - net_approved

    status = "VERIFIED OK"
    if outage_penalty > 0 or total_ded > 50.0 or (tenancy != claimed_tenancy):
        status = "DISPUTED / DEDUCTION"
    elif remark != "" and any(w in remark.lower() for w in ["theft", "down", "deleted", "fault", "off"]):
        status = "FIELD WARNING"

    rows_data.append({
        "invoice_no": inv_str,
        "global_id": gid_str,
        "vendor_id": vendor_id,
        "bsnl_id": bsnl_id,
        "site_name": site_name,
        "ssa": ssa,
        "eqpt_type": eqpt_type,
        "bts_eqpt": bts_eqpt,
        "claimed_tenancy": claimed_tenancy,
        "working_tenancy": tenancy,
        "tenancy_diff": tenancy - claimed_tenancy,
        "ssa_type": ssa_type,
        "claimed_city": claimed_city,
        "dg_cap": dg_cap,
        "bts_rental_type": bts_rental_type,
        "dg_run": dg_run,
        "dg_cph": dg_cph,
        "diesel_rate": diesel_rate,
        "diesel_chg_day": diesel_chg_day,
        "eb_hrs": eb_hrs,
        "eb_rate": eb_rate,
        "eb_cph": eb_cph,
        "eb_chg_day": eb_chg_day,
        "total_chg_day": total_chg_day,
        "calculated_month": total_chg_month,
        "claimed_amount": claimed_amount,
        "passed_base": passed_base,
        "outage_min": outage_min,
        "allowed_min": ALLOWED_OUTAGE_MINS if outage_min > 0 else 0.0,
        "excess_min": excess_mins,
        "outage_penalty": outage_penalty,
        "net_approved": net_approved,
        "total_deduction": total_ded,
        "audit_status": status,
        "outage_note": outage_note,
        "remark": remark
    })

df_all = pd.DataFrame(rows_data)
print(f"Total sites ingested: {len(df_all)}")

# 2. Build Multi-Tab Styled Workbook
wb = openpyxl.Workbook()

# Style Palette
navy_header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
sub_header_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
green_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
yellow_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
red_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC")
)

font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
font_head = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
font_bold = Font(name="Calibri", size=10, bold=True)
font_cell = Font(name="Calibri", size=10)

def autofit(ws, min_col=1, max_col=None):
    ws.views.sheetView[0].showGridLines = True
    cols_to_check = ws.columns if max_col is None else list(ws.columns)[min_col-1:max_col]
    for col in cols_to_check:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row in [1, 2]:
                continue
            val_s = str(cell.value or "")
            if len(val_s) > max_len:
                max_len = len(val_s)
        ws.column_dimensions[col_letter].width = max(11, min(max_len + 3, 40))

# ----------------- SHEET 1: EXECUTIVE SUMMARY -----------------
ws_sum = wb.active
ws_sum.title = "Executive Summary (5 Invoices)"

ws_sum.merge_cells("A1:K1")
ws_sum["A1"] = "BHARAT SANCHAR NIGAM LIMITED - IP BILLING & SLA RECONCILIATION AUDIT"
ws_sum["A1"].font = font_title
ws_sum["A1"].fill = navy_header_fill
ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_sum.row_dimensions[1].height = 40

ws_sum.merge_cells("A2:K2")
ws_sum["A2"] = "Summary of 5 Monthly Invoices from M/s Elevar Digitel (ATC) with Outage Certificate Verification (Feb-2026)"
ws_sum["A2"].font = Font(name="Calibri", size=11, italic=True)
ws_sum["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws_sum.row_dimensions[2].height = 25

headers_summary = [
    "S.No", "Invoice No", "PDF Pages", "Sites Count",
    "Claimed Excl. GST (₹)", "BSNL Base Passed (₹)",
    "Outage Penalty (₹)", "Net Approved Excl. GST (₹)",
    "Total Deduction (₹)", "GST @18% on Passed (₹)", "Net Payable With GST (₹)"
]

ws_sum.append([])
ws_sum.append(headers_summary)
ws_sum.row_dimensions[4].height = 28
for c in range(1, len(headers_summary) + 1):
    cell = ws_sum.cell(row=4, column=c)
    cell.font = font_head
    cell.fill = navy_header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

inv_order = [
    ("607239111046515", "Pages 1-5"),
    ("607239111046516", "Pages 6-9"),
    ("607239111046519", "Pages 10-14"),
    ("607239111046520", "Pages 15-19"),
    ("607239111046533", "Pages 20-24")
]

s_no = 1
for inv_num, pdf_pg in inv_order:
    sub = df_all[df_all["invoice_no"] == inv_num]
    cnt = len(sub)
    c_amt = sub["claimed_amount"].sum()
    b_pass = sub["passed_base"].sum()
    pen_amt = sub["outage_penalty"].sum()
    net_app = sub["net_approved"].sum()
    ded_amt = sub["total_deduction"].sum()
    gst_pass = net_app * 0.18
    tot_pass = net_app + gst_pass

    ws_sum.append([
        s_no, inv_num, pdf_pg, cnt,
        c_amt, b_pass, pen_amt, net_app, ded_amt, gst_pass, tot_pass
    ])
    r = ws_sum.max_row
    ws_sum.row_dimensions[r].height = 22
    for c in range(1, 12):
        cell = ws_sum.cell(row=r, column=c)
        cell.font = font_cell
        cell.border = thin_border
        if c in [1, 2, 3, 4]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "#,##0.00"
    s_no += 1

# Total row
total_c = df_all["claimed_amount"].sum()
total_b = df_all["passed_base"].sum()
total_p = df_all["outage_penalty"].sum()
total_n = df_all["net_approved"].sum()
total_d = df_all["total_deduction"].sum()
total_g = total_n * 0.18
total_w = total_n + total_g

ws_sum.append([
    "TOTAL", "ALL 5 INVOICES", "Pages 1-24", len(df_all),
    total_c, total_b, total_p, total_n, total_d, total_g, total_w
])
tot_row = ws_sum.max_row
ws_sum.row_dimensions[tot_row].height = 26
for c in range(1, 12):
    cell = ws_sum.cell(row=tot_row, column=c)
    cell.font = font_bold
    cell.fill = sub_header_fill
    cell.border = thin_border
    if c in [1, 2, 3, 4]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
    else:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "#,##0.00"

# Summary Statistics Block
ws_sum.append([])
ws_sum.append(["AUDIT FINDINGS & RECOVERY SUMMARY", "", "", ""])
ws_sum.merge_cells(f"A{ws_sum.max_row}:D{ws_sum.max_row}")
ws_sum.cell(row=ws_sum.max_row, column=1).font = font_bold
ws_sum.cell(row=ws_sum.max_row, column=1).fill = sub_header_fill

stat_rows = [
    ("Total Sites Audited Across 5 Invoices", f"{len(df_all)} Sites"),
    ("Sites With Verified Outage SLA Penalties", f"{(df_all['outage_penalty'] > 0).sum()} Sites"),
    ("Total SLA Outage Penalty Deduction", f"₹{total_p:,.2f}"),
    ("Total Rate & Energy Discrepancy Deductions", f"₹{(total_d - total_p):,.2f}"),
    ("Total Proposed Deduction / Debit Note (Excl. GST)", f"₹{total_d:,.2f}"),
    ("Total Proposed Deduction / Debit Note (Incl. 18% GST)", f"₹{(total_d * 1.18):,.2f}"),
    ("Net BSNL Approved Payable (Excl. GST)", f"₹{total_n:,.2f}"),
    ("Net BSNL Approved Payable (Incl. 18% GST)", f"₹{total_w:,.2f}")
]

for desc, val in stat_rows:
    ws_sum.append([desc, "", val, ""])
    r = ws_sum.max_row
    ws_sum.merge_cells(f"A{r}:B{r}")
    ws_sum.merge_cells(f"C{r}:D{r}")
    ws_sum.cell(row=r, column=1).font = font_cell
    ws_sum.cell(row=r, column=1).border = thin_border
    ws_sum.cell(row=r, column=3).font = font_bold
    ws_sum.cell(row=r, column=3).border = thin_border
    ws_sum.cell(row=r, column=3).alignment = Alignment(horizontal="right", vertical="center")

autofit(ws_sum)

# ----------------- SHEETS 2-6: INVOICE-BY-INVOICE SHEETS -----------------
site_headers = [
    "S.No", "Global ID", "Vendor Site ID", "BSNL Site ID", "Site Name", "SSA",
    "City Class", "BTS Type", "Claimed Tenancy", "Verified Tenancy", "Tenancy Diff",
    "DG Cap", "DG Run/Day", "DG CPH", "Diesel Rate", "EB Hrs/Day", "EB Rate", "EB CPH",
    "Claimed Energy (₹)", "BSNL Calculated (₹)", "Base Passed (₹)",
    "Outage Mins (Cert.)", "Allowed Mins (40.32m)", "Excess Mins", "Outage Penalty (₹)",
    "Net Approved (₹)", "Total Deduction (₹)", "Audit Status", "Field Officer Remarks & SLA Note"
]

inv_sheet_titles = {
    "607239111046515": "Inv 1 (607239111046515)",
    "607239111046516": "Inv 2 (607239111046516)",
    "607239111046519": "Inv 3 (607239111046519)",
    "607239111046520": "Inv 4 (607239111046520)",
    "607239111046533": "Inv 5 (607239111046533)"
}

for inv_num, sheet_title in inv_sheet_titles.items():
    ws_inv = wb.create_sheet(title=sheet_title[:31])
    sub_df = df_all[df_all["invoice_no"] == inv_num].reset_index(drop=True)

    # Title block
    ws_inv.merge_cells("A1:K1")
    ws_inv["A1"] = f"INVOICE NO: {inv_num} - SITE-WISE BILLING & SLA RECONCILIATION ({len(sub_df)} SITES)"
    ws_inv["A1"].font = font_title
    ws_inv["A1"].fill = navy_header_fill
    ws_inv["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_inv.row_dimensions[1].height = 35

    ws_inv.append([])
    ws_inv.append(site_headers)
    h_row = 3
    ws_inv.row_dimensions[h_row].height = 28
    for c in range(1, len(site_headers) + 1):
        cell = ws_inv.cell(row=h_row, column=c)
        cell.font = font_head
        cell.fill = navy_header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    idx = 1
    for _, row in sub_df.iterrows():
        note_str = row["outage_note"]
        if row["remark"]:
            note_str += f" | Remark: {row['remark']}"

        ws_inv.append([
            idx,
            row["global_id"],
            row["vendor_id"],
            row["bsnl_id"],
            row["site_name"],
            row["ssa"],
            row["ssa_type"],
            row["bts_rental_type"],
            row["claimed_tenancy"],
            row["working_tenancy"],
            row["tenancy_diff"],
            row["dg_cap"],
            row["dg_run"],
            row["dg_cph"],
            row["diesel_rate"],
            row["eb_hrs"],
            row["eb_rate"],
            row["eb_cph"],
            row["claimed_amount"],
            row["calculated_month"],
            row["passed_base"],
            row["outage_min"],
            row["allowed_min"],
            row["excess_min"],
            row["outage_penalty"],
            row["net_approved"],
            row["total_deduction"],
            row["audit_status"],
            note_str
        ])
        curr_r = ws_inv.max_row
        ws_inv.row_dimensions[curr_r].height = 20

        # Row styling
        fill = green_fill
        if row["audit_status"] == "DISPUTED / DEDUCTION":
            fill = red_fill
        elif row["audit_status"] == "FIELD WARNING":
            fill = yellow_fill

        for c in range(1, len(site_headers) + 1):
            cell = ws_inv.cell(row=curr_r, column=c)
            cell.font = font_cell
            cell.border = thin_border
            # Numbers formatting
            if c in [19, 20, 21, 25, 26, 27]:
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c in [1, 2, 6, 7, 8, 9, 10, 11, 22, 23, 24]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Status highlight
            if c == 28:
                cell.fill = fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = font_bold
            if c == 25 and row["outage_penalty"] > 0:
                cell.fill = red_fill
                cell.font = font_bold

        idx += 1

    # Invoice Total Row
    tot_c = sub_df["claimed_amount"].sum()
    tot_p = sub_df["passed_base"].sum()
    tot_pen = sub_df["outage_penalty"].sum()
    tot_net = sub_df["net_approved"].sum()
    tot_ded = sub_df["total_deduction"].sum()

    ws_inv.append([
        "TOTAL", f"{len(sub_df)} SITES", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
        tot_c, "", tot_p, "", "", "", tot_pen, tot_net, tot_ded, "", ""
    ])
    inv_tot_r = ws_inv.max_row
    ws_inv.row_dimensions[inv_tot_r].height = 24
    for c in range(1, len(site_headers) + 1):
        cell = ws_inv.cell(row=inv_tot_r, column=c)
        cell.font = font_bold
        cell.fill = sub_header_fill
        cell.border = thin_border
        if c in [19, 21, 25, 26, 27]:
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right", vertical="center")

    autofit(ws_inv)

# ----------------- SHEET 7: OUTAGE & SLA AUDIT -----------------
ws_out = wb.create_sheet(title="Outage & SLA Audit")
ws_out.merge_cells("A1:K1")
ws_out["A1"] = "BSNL SLA OUTAGE AUDIT & CERTIFICATE RECONCILIATION - FEB-2026"
ws_out["A1"].font = font_title
ws_out["A1"].fill = navy_header_fill
ws_out["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_out.row_dimensions[1].height = 35

out_headers = [
    "S.No", "Invoice No", "Global ID", "Vendor ID", "Site Name", "SSA",
    "Outage Mins (Cert.)", "Permissible Mins (40.32m)", "Excess Mins", "Excess Hours",
    "Calculated Penalty (₹)", "Penalty Cap (₹)", "Final Penalty Levied (₹)",
    "Officer Notes / Outage Sign-off Certificate Details"
]

ws_out.append([])
ws_out.append(out_headers)
ws_out.row_dimensions[3].height = 28
for c in range(1, len(out_headers) + 1):
    cell = ws_out.cell(row=3, column=c)
    cell.font = font_head
    cell.fill = navy_header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

outage_df = df_all[df_all["outage_min"] > 0].sort_values(by=["invoice_no", "global_id"]).reset_index(drop=True)
o_idx = 1
for _, row in outage_df.iterrows():
    excess_h = row["excess_min"] / 60.0
    raw_pen = excess_h * PENALTY_RATE_PER_HR

    ws_out.append([
        o_idx,
        row["invoice_no"],
        row["global_id"],
        row["vendor_id"],
        row["site_name"],
        row["ssa"],
        row["outage_min"],
        ALLOWED_OUTAGE_MINS,
        row["excess_min"],
        round(excess_h, 2),
        raw_pen,
        PENALTY_CAP,
        row["outage_penalty"],
        row["outage_note"] + (f" | {row['remark']}" if row['remark'] else "")
    ])
    curr_r = ws_out.max_row
    ws_out.row_dimensions[curr_r].height = 20
    for c in range(1, len(out_headers) + 1):
        cell = ws_out.cell(row=curr_r, column=c)
        cell.font = font_cell
        cell.border = thin_border
        if c in [11, 12, 13]:
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if c == 13:
                cell.fill = red_fill
                cell.font = font_bold
        elif c in [1, 2, 3, 6, 7, 8, 9, 10]:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    o_idx += 1

# Outage Total Row
ws_out.append([
    "TOTAL", f"{len(outage_df)} SITES", "", "", "", "",
    outage_df["outage_min"].sum(), "", outage_df["excess_min"].sum(), "",
    "", "", outage_df["outage_penalty"].sum(), ""
])
o_tot_r = ws_out.max_row
ws_out.row_dimensions[o_tot_r].height = 24
for c in range(1, len(out_headers) + 1):
    cell = ws_out.cell(row=o_tot_r, column=c)
    cell.font = font_bold
    cell.fill = sub_header_fill
    cell.border = thin_border
    if c == 13:
        cell.number_format = "#,##0.00"
        cell.alignment = Alignment(horizontal="right", vertical="center")

autofit(ws_out)

# ----------------- SHEET 8: DEBIT NOTE / DISPUTE ANNEXURE -----------------
ws_deb = wb.create_sheet(title="Vendor Debit Note Annexure")
ws_deb.merge_cells("A1:I1")
ws_deb["A1"] = "BSNL PROPOSED DEBIT NOTE / DISPUTE ANNEXURE - ELEVAR DIGITEL (ATC)"
ws_deb["A1"].font = font_title
ws_deb["A1"].fill = navy_header_fill
ws_deb["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_deb.row_dimensions[1].height = 35

deb_headers = [
    "S.No", "Invoice No", "Global ID", "Site Name", "SSA",
    "Claimed Energy (₹)", "BSNL Approved (₹)", "Outage Penalty (₹)",
    "Proposed Debit / Deduction (₹)", "Dispute / Deduction Reason"
]

ws_deb.append([])
ws_deb.append(deb_headers)
ws_deb.row_dimensions[3].height = 28
for c in range(1, len(deb_headers) + 1):
    cell = ws_deb.cell(row=3, column=c)
    cell.font = font_head
    cell.fill = navy_header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

disputed_df = df_all[df_all["total_deduction"] > 1.0].sort_values(by="total_deduction", ascending=False).reset_index(drop=True)
d_idx = 1
for _, row in disputed_df.iterrows():
    reasons = []
    if row["outage_penalty"] > 0:
        reasons.append(f"Outage SLA Penalty (₹{row['outage_penalty']:,.2f})")
    if row["claimed_amount"] > row["passed_base"]:
        reasons.append(f"HPC Energy norm excess claim (₹{row['claimed_amount'] - row['passed_base']:,.2f})")
    if row["tenancy_diff"] != 0:
        reasons.append(f"Tenancy mismatch (Claimed {row['claimed_tenancy']} vs Field {row['working_tenancy']})")
    if row["remark"]:
        reasons.append(row["remark"])

    reason_str = "; ".join(reasons)

    ws_deb.append([
        d_idx,
        row["invoice_no"],
        row["global_id"],
        row["site_name"],
        row["ssa"],
        row["claimed_amount"],
        row["net_approved"],
        row["outage_penalty"],
        row["total_deduction"],
        reason_str
    ])
    curr_r = ws_deb.max_row
    ws_deb.row_dimensions[curr_r].height = 20
    for c in range(1, len(deb_headers) + 1):
        cell = ws_deb.cell(row=curr_r, column=c)
        cell.font = font_cell
        cell.border = thin_border
        if c in [6, 7, 8, 9]:
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if c == 9:
                cell.font = font_bold
                cell.fill = red_fill
        elif c in [1, 2, 3, 5]:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    d_idx += 1

# Debit Total Row
ws_deb.append([
    "TOTAL", f"{len(disputed_df)} SITES", "", "", "",
    disputed_df["claimed_amount"].sum(), disputed_df["net_approved"].sum(),
    disputed_df["outage_penalty"].sum(), disputed_df["total_deduction"].sum(), ""
])
deb_tot_r = ws_deb.max_row
ws_deb.row_dimensions[deb_tot_r].height = 24
for c in range(1, len(deb_headers) + 1):
    cell = ws_deb.cell(row=deb_tot_r, column=c)
    cell.font = font_bold
    cell.fill = sub_header_fill
    cell.border = thin_border
    if c in [6, 7, 8, 9]:
        cell.number_format = "#,##0.00"
        cell.alignment = Alignment(horizontal="right", vertical="center")

autofit(ws_deb)

# Save output file
output_file = "BSNL_5_Invoices_Reconciliation_Output.xlsx"
wb.save(output_file)
print(f"Workbook saved successfully as '{output_file}' ({os.path.getsize(output_file):,} bytes)!")
