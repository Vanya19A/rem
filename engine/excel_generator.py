"""
Professional Excel Generator for BSNL Telecom Infrastructure Reconciled Reports
Generates styled, 4-tab Excel workbooks matching official BSNL audit standards.
"""

import io
from datetime import datetime
from typing import Dict, Any, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class ExcelReportGenerator:
    def __init__(self):
        # Color palette
        self.bsnl_blue = "003366"      # Deep Navy Blue
        self.bsnl_gold = "CC9900"      # BSNL Gold/Amber
        self.header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        self.sub_header_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        
        # Status fills
        self.dispute_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")   # Light red
        self.warning_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")   # Light yellow
        self.ok_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")        # Light green

        # Borders & Fonts
        thin = Side(border_style="thin", color="D3D3D3")
        self.border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
        self.font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        self.font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        self.font_bold = Font(name="Calibri", size=11, bold=True)
        self.font_regular = Font(name="Calibri", size=10)

    def generate_workbook(
        self,
        reconciled_df: pd.DataFrame,
        summary_metrics: Dict[str, Any],
        vendor_name: str = "Elevar Digitel (ATC)",
        billing_period: str = "Feb-2026"
    ) -> bytes:
        """Create the complete multi-tab audit workbook in memory."""
        wb = openpyxl.Workbook()

        # Sheet 1: Executive Summary & MIRO
        ws_sum = wb.active
        ws_sum.title = "Executive Summary & MIRO"
        self._build_summary_sheet(ws_sum, summary_metrics, vendor_name, billing_period)

        # Sheet 2: Site-wise Reconciliation
        ws_sites = wb.create_sheet(title="Site-Wise Reconciliation")
        self._build_sites_sheet(ws_sites, reconciled_df)

        # Sheet 3: Outage & SLA Penalty Details
        ws_outage = wb.create_sheet(title="Outage & SLA Audit")
        self._build_outage_sheet(ws_outage, reconciled_df)

        # Sheet 4: Vendor Dispute / Debit Note
        ws_dispute = wb.create_sheet(title="Vendor Debit Note Annexure")
        self._build_dispute_sheet(ws_dispute, reconciled_df, vendor_name, billing_period)

        # Save to buffer
        out_buf = io.BytesIO()
        wb.save(out_buf)
        return out_buf.getvalue()

    def _build_summary_sheet(self, ws, metrics: Dict[str, Any], vendor: str, period: str):
        """Construct Executive Summary and MIRO entry sheet."""
        ws.views.sheetView[0].showGridLines = True
        
        # Title Block
        ws.merge_cells("A1:G1")
        top_cell = ws["A1"]
        top_cell.value = f"BHARAT SANCHAR NIGAM LIMITED - IP BILLING & SLA RECONCILIATION AUDIT"
        top_cell.font = self.font_title
        top_cell.fill = self.header_fill
        top_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40

        ws.merge_cells("A2:G2")
        ws["A2"].value = f"Vendor: {vendor}  |  Billing Period: {period}  |  Generated on: {datetime.now().strftime('%d-%b-%Y %H:%M')}"
        ws["A2"].font = Font(name="Calibri", size=11, italic=True)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 25

        # Key Metrics Table
        headers = ["Parameter", "Amount (Excl. GST)", "GST @ 18%", "Total Amount (Incl. GST)", "Remarks / Clause"]
        ws.append([])
        ws.append(headers)
        h_row = 4
        ws.row_dimensions[h_row].height = 28
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=h_row, column=col_idx)
            cell.font = self.font_header
            cell.fill = self.header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        data_rows = [
            ("Total Claimed Amount (Vendor Invoice)", metrics["total_claimed_before_tax"], metrics["total_claimed_before_tax"] * 0.18, metrics["total_claimed_with_gst"], "As submitted in vendor monthly invoice"),
            ("Outage SLA Penalties (Deduction)", metrics["total_outage_penalty"], metrics["total_outage_penalty"] * 0.18, metrics["total_outage_penalty"] * 1.18, "Deduction as per MSA Uptime SLA Clause 7.0 (₹500/hr)"),
            ("Other Rate / Tenancy Overbilling Deductions", metrics["total_deduction_before_tax"] - metrics["total_outage_penalty"], (metrics["total_deduction_before_tax"] - metrics["total_outage_penalty"]) * 0.18, (metrics["total_deduction_before_tax"] - metrics["total_outage_penalty"]) * 1.18, "Rate card ceiling / Tenancy / Decommissioned sites"),
            ("Total Deductions / Disputed Amount", metrics["total_deduction_before_tax"], metrics["total_deduction_before_tax"] * 0.18, metrics["total_deduction_with_gst"], "Net recovery / Credit note to be issued by Vendor"),
            ("Net Payable Amount (BSNL Approved)", metrics["total_passed_before_tax"], metrics["total_passed_before_tax"] * 0.18, metrics["total_passed_with_gst"], "Recommended for MIRO / Payment approval in ERP"),
        ]

        for row_data in data_rows:
            ws.append([
                row_data[0],
                f"₹{row_data[1]:,.2f}",
                f"₹{row_data[2]:,.2f}",
                f"₹{row_data[3]:,.2f}",
                row_data[4]
            ])
            r = ws.max_row
            ws.row_dimensions[r].height = 24
            for c in range(1, 6):
                cell = ws.cell(row=r, column=c)
                cell.font = self.font_regular
                cell.border = self.border_thin
                if c in [2, 3, 4]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    if "Net Payable" in row_data[0]:
                        cell.font = Font(name="Calibri", size=11, bold=True, color="006600")
                    elif "Total Deductions" in row_data[0]:
                        cell.font = Font(name="Calibri", size=11, bold=True, color="990000")

        # Site counts overview
        ws.append([])
        ws.append(["Site Audit Status", "Count of Sites", "% of Total Sites"])
        s_row = ws.max_row
        for c in range(1, 4):
            ws.cell(row=s_row, column=c).font = self.font_header
            ws.cell(row=s_row, column=c).fill = self.sub_header_fill
            ws.cell(row=s_row, column=c).alignment = Alignment(horizontal="center", vertical="center")

        total = metrics["total_sites"]
        status_rows = [
            ("Verified Sites (No Discrepancy)", metrics["verified_sites_count"], f"{(metrics['verified_sites_count']/max(1, total))*100:.1f}%"),
            ("Field Warning Sites (Remarks / Hardware Faults)", metrics["warning_sites_count"], f"{(metrics['warning_sites_count']/max(1, total))*100:.1f}%"),
            ("Disputed / Deduction Sites (Overbilling / Penalty)", metrics["disputed_sites_count"], f"{(metrics['disputed_sites_count']/max(1, total))*100:.1f}%"),
            ("Total Sites Audited", total, "100.0%")
        ]

        for s_data in status_rows:
            ws.append([s_data[0], s_data[1], s_data[2]])
            r = ws.max_row
            for c in range(1, 4):
                ws.cell(row=r, column=c).font = self.font_regular
                ws.cell(row=r, column=c).border = self.border_thin
                if c >= 2:
                    ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")

        # Auto-adjust column widths
        self._autofit_columns(ws)

    def _build_sites_sheet(self, ws, df: pd.DataFrame):
        """Build Full Site-by-Site Reconciliation Table."""
        ws.views.sheetView[0].showGridLines = True

        export_cols = [
            "global_id", "vendor_site_id", "bsnl_site_id", "site_name", "ssa",
            "working_tenancy", "claimed_tenancy", "tenancy_diff",
            "outage_minutes", "outage_penalty_rs",
            "claimed_amount", "calculated_energy_amount", "passed_base_amount",
            "net_payable_amount", "total_deduction", "audit_status", "field_remark"
        ]
        
        # Filter existing columns
        cols = [c for c in export_cols if c in df.columns]
        headers = [c.replace("_", " ").title() for c in cols]
        
        ws.append(headers)
        ws.row_dimensions[1].height = 28
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = self.font_header
            cell.fill = self.header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for _, row in df.iterrows():
            row_vals = []
            for col in cols:
                val = row[col]
                if pd.isna(val):
                    row_vals.append("")
                elif isinstance(val, (int, float)):
                    if "amount" in col or "penalty" in col or "deduction" in col:
                        row_vals.append(round(float(val), 2))
                    else:
                        row_vals.append(val)
                else:
                    row_vals.append(str(val))
            
            ws.append(row_vals)
            r = ws.max_row
            ws.row_dimensions[r].height = 20

            # Formatting
            status = str(row.get("audit_status", ""))
            fill = self.ok_fill
            if "DISPUTED" in status:
                fill = self.dispute_fill
            elif "WARNING" in status:
                fill = self.warning_fill

            for c in range(1, len(cols) + 1):
                cell = ws.cell(row=r, column=c)
                cell.font = self.font_regular
                cell.border = self.border_thin
                
                # Highlight status column
                if cols[c-1] == "audit_status":
                    cell.fill = fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        self._autofit_columns(ws)

    def _build_outage_sheet(self, ws, df: pd.DataFrame):
        """Build Outage & SLA Penalty details sheet."""
        ws.views.sheetView[0].showGridLines = True

        outage_sites = df[df["outage_minutes"] > 0].copy()
        if outage_sites.empty:
            outage_sites = df.head(10).copy()

        cols = ["global_id", "vendor_site_id", "site_name", "ssa", "outage_minutes", "excess_outage_minutes", "outage_penalty_rs", "outage_audit_note", "field_remark"]
        existing_cols = [c for c in cols if c in outage_sites.columns]
        headers = [c.replace("_", " ").title() for c in existing_cols]

        ws.append(headers)
        ws.row_dimensions[1].height = 28
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = self.font_header
            cell.fill = self.header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for _, row in outage_sites.iterrows():
            row_vals = [row.get(c, "") for c in existing_cols]
            ws.append(row_vals)
            r = ws.max_row
            ws.row_dimensions[r].height = 20
            for c in range(1, len(existing_cols) + 1):
                cell = ws.cell(row=r, column=c)
                cell.font = self.font_regular
                cell.border = self.border_thin
                if existing_cols[c-1] == "outage_penalty_rs" and float(row.get("outage_penalty_rs", 0)) > 0:
                    cell.fill = self.dispute_fill
                    cell.font = Font(name="Calibri", size=10, bold=True, color="990000")

        self._autofit_columns(ws)

    def _build_dispute_sheet(self, ws, df: pd.DataFrame, vendor: str, period: str):
        """Build Debit Note annexure sheet of all disputed / penalized sites."""
        ws.views.sheetView[0].showGridLines = True

        disputed = df[df["total_deduction"] > 0.01].copy()
        if disputed.empty:
            disputed = df.head(5).copy()

        # Header Block
        ws.merge_cells("A1:F1")
        ws["A1"].value = f"BSNL DISPUTE MATRIX & PROPOSED DEBIT NOTE - {vendor.upper()}"
        ws["A1"].font = self.font_title
        ws["A1"].fill = self.header_fill
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 35

        ws.merge_cells("A2:F2")
        ws["A2"].value = f"Billing Period: {period}  |  Clause Reference: Master Service Agreement & SLA Schedule 2"
        ws["A2"].font = Font(name="Calibri", size=10, italic=True)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

        ws.append([])
        cols = ["global_id", "vendor_site_id", "site_name", "claimed_amount", "net_payable_amount", "total_deduction", "outage_penalty_rs", "field_remark"]
        existing_cols = [c for c in cols if c in disputed.columns]
        headers = [c.replace("_", " ").title() for c in existing_cols]

        ws.append(headers)
        h_row = ws.max_row
        ws.row_dimensions[h_row].height = 26
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=h_row, column=c)
            cell.font = self.font_header
            cell.fill = self.header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for _, row in disputed.iterrows():
            row_vals = [row.get(c, "") for c in existing_cols]
            ws.append(row_vals)
            r = ws.max_row
            ws.row_dimensions[r].height = 20
            for c in range(1, len(existing_cols) + 1):
                cell = ws.cell(row=r, column=c)
                cell.font = self.font_regular
                cell.border = self.border_thin

        self._autofit_columns(ws)

    def _autofit_columns(self, ws):
        """Auto-adjust column widths based on content."""
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                # avoid inspecting merged cells title
                if cell.row in [1, 2]:
                    continue
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(12, min(max_len + 4, 45))
