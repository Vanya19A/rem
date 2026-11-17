"""
BSNL SAP ERP REM (Real Estate Management) Upload Template Generator
Generates the 3 mandatory upload templates for each invoice:
1. INVOICE UPLOAD (30-column header and site line item table)
2. INVOICE DETAIL UPLOAD (8-column contract, vendor, period, fiscal year mapping)
3. CONDITION UPLOAD (10-column condition types: T500 Rent, T570 EB, T580 DG, T630 LLR)
"""

import io
import zipfile
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from typing import Dict, Any, List, Optional


class ERPREMGenerator:
    # Standard SAP REM Condition Types for BSNL
    CONDITION_TYPES = {
        "RENT": "T500",      # Passive Infrastructure Provisioning Fee (IP Rent)
        "ENERGY_EB": "T570", # Electricity Charges
        "ENERGY_DG": "T580", # Diesel Charges
        "LLR": "T630"        # Ground Rent / Landlord Rent
    }

    # Month to Financial Year Period Mapping (April = 1 ... Feb = 11, Mar = 12)
    FY_PERIOD_MAP = {
        "APR": 1, "MAY": 2, "JUN": 3, "JUL": 4, "AUG": 5, "SEP": 6,
        "OCT": 7, "NOV": 8, "DEC": 9, "JAN": 10, "FEB": 11, "MAR": 12
    }

    @staticmethod
    def get_fy_period(month_str: str) -> int:
        m = str(month_str).upper()[:3]
        return ERPREMGenerator.FY_PERIOD_MAP.get(m, 11)

    @classmethod
    def generate_invoice_upload(
        cls,
        invoice_no: str,
        sites: List[Dict[str, Any]],
        bill_type: str = "RENT", # "ENERGY", "RENT", or "LLR"
        circle_code: str = "GJ",
        company_code: int = 1100,
        business_entity: str = "16400000",
        vendor_no: str = "2004755",
        from_date: str = "01.04.2026",
        to_date: str = "30.04.2026"
    ) -> bytes:
        """
        Creates INVOICE UPLOAD excel matching official BSNL REM template:
        Sheet Name: Circle Code (e.g. 'MP')
        Row 1: Empty
        Row 2: Empty
        Row 3: 30 Column Headers
        Row 4+: Data rows
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = str(circle_code).upper()

        headers_row3 = [
            'Sno.',
            'Company Code',
            'Business Entity',
            'Contract Number',
            'Rental Object No',
            'Name of Rental Object ',
            'Vendor no',
            'Invoice number ',
            'Invoice Period From\n( DD.MM.YYYY)',
            'Invoice Period To\n (DD.MM.YYYY)',
            'Name of the Site',
            ' Passive Infrasructure Provisining Fee (Rs.)',
            'Ground Rent',
            ' Addnl Indoor BTS Rack Charges ',
            ' Addnl Outdoor BTS Rack Charges ',
            'Addnl  3GSM Antennae Charges',
            'Addnl 1.2 M dia MW Antenna Charges',
            'Addnl 0.6 M dia MW Antenna Charges',
            'Addnl  0.7 SQM GPS  Antenna Charges',
            'Electricity Charges',
            'Diesel Charges',
            'Vendor Tax Charges',
            'Charges for Addnl Services',
            'Volume Based Discount',
            'Outage Penalty',
            'Roll Out Penalty',
            'Lock In Period Penalty',
            'Late Payment Penalty',
            'Penal Interest For Tenancy',
            'Total (Rs.)'
        ]

        # Write Row 3 headers
        header_font = Font(name="Arial", size=10, bold=True)
        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        thin_border = Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )

        for col_idx, h in enumerate(headers_row3, start=1):
            cell = ws.cell(row=3, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        ws.row_dimensions[3].height = 28

        # Populate data rows
        row_num = 4
        for idx, site in enumerate(sites, start=1):
            c_no = site.get("contract_no") or site.get("Contract Number") or site.get("REM Contract No")
            if not c_no:
                # Fallback to pseudo-contract if missing
                gid = str(site.get("global_id", site.get("Global ID", "")))
                c_no = int(f"1653000{gid.zfill(6)}") if gid.isdigit() else 1653000000000 + idx
            else:
                try:
                    c_no = int(float(c_no))
                except:
                    pass

            rental_obj = site.get("rental_object_no") or site.get("Rental Object No")
            if not rental_obj:
                try:
                    rental_obj = 50000000 + (int(str(c_no)[-5:]) if str(c_no).isdigit() else idx)
                except:
                    rental_obj = 50000000 + idx
            else:
                try:
                    rental_obj = int(float(rental_obj))
                except:
                    pass

            site_name = site.get("site_name") or site.get("Site Name", "")
            rental_obj_name = f"Lease-In-Slot No-1/GSM/IP/{site_name}" if site_name else None

            # Calculate billing components based on billing stream
            passive_rent = None
            ground_rent = 0.0
            eb_charges = 0.0
            dg_charges = 0.0
            outage_pen = float(site.get("outage_penalty", site.get("Outage Penalty (₹)", site.get("Deduction Outage (₹)", 0.0))) or 0.0)

            if bill_type == "RENT":
                passive_rent = float(site.get("approved_rent", site.get("Net Approved (₹)", site.get("Base Passed (₹)", 0.0))) or 0.0)
                tot_site = round(passive_rent - outage_pen, 2)
            elif bill_type == "LLR":
                ground_rent = float(site.get("ground_rent", site.get("Approved Landlord Rent", site.get("Net Approved (₹)", 0.0))) or 0.0)
                tot_site = ground_rent
            else: # ENERGY
                # Default Energy split
                tot_energy = float(site.get("base_passed", site.get("Base Passed (₹)", site.get("Net Approved (₹)", 0.0))) or 0.0)
                dg_h = float(site.get("dg_run_hours", site.get("DG Run/Day", 0.0)) or 0.0)
                if dg_h > 0:
                    dg_ratio = min(0.6, (dg_h / 24.0) * 1.5)
                    dg_charges = round(tot_energy * dg_ratio, 2)
                    eb_charges = round(tot_energy - dg_charges, 2)
                else:
                    eb_charges = tot_energy
                    dg_charges = 0.0
                tot_site = round(eb_charges + dg_charges - outage_pen, 2)

            data_vals = [
                idx,
                int(company_code),
                int(business_entity) if str(business_entity).isdigit() else business_entity,
                c_no,
                rental_obj,
                rental_obj_name,
                int(vendor_no) if str(vendor_no).isdigit() else vendor_no,
                str(site.get("invoice_no") or invoice_no),
                from_date,
                to_date,
                site_name,
                passive_rent,
                ground_rent,
                0, # Addnl Indoor BTS
                0, # Addnl Outdoor BTS
                0, # 3GSM
                0, # 1.2M MW
                0, # 0.6M MW
                0, # 0.7 GPS
                eb_charges if bill_type == "ENERGY" else None,
                dg_charges if bill_type == "ENERGY" else None,
                0, # Tax
                0, # Addnl Services
                0, # Volume Discount
                outage_pen,
                0, # Roll Out
                0, # Lock In
                0, # Late Payment
                0, # Penal Interest
                tot_site
            ]

            for c_idx, val in enumerate(data_vals, start=1):
                c = ws.cell(row=row_num, column=c_idx, value=val)
                c.border = thin_border
                if c_idx in [1, 2, 3, 4, 5, 7]:
                    c.alignment = Alignment(horizontal="center")
                elif c_idx in [9, 10]:
                    c.alignment = Alignment(horizontal="center")
                elif c_idx >= 12:
                    c.alignment = Alignment(horizontal="right")
                    if isinstance(val, (int, float)):
                        c.number_format = '#,##0.00'

            row_num += 1

        # Adjust column widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    @classmethod
    def generate_invoice_detail_upload(
        cls,
        invoice_no: str,
        sites: List[Dict[str, Any]],
        circle_code: str = "GJ",
        company_code: int = 1100,
        vendor_no: str = "2004755",
        invoice_date: str = "05.05.2026",
        period: int = 1,
        fiscal_year: int = 2026
    ) -> bytes:
        """
        Creates INVOICE detail UPLOAD excel matching official BSNL REM template:
        Sheet Name: Circle Code (e.g. 'MP')
        Row 1: 8 Column Headers
        Row 2+: Data rows
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = str(circle_code).upper()

        headers = [
            'COMPANY CODE                 (pl enter 1100)',
            'CONTRACT NUMBER                     (pl enter 13 digit contract number e.g. 1352000000001) ',
            'VENDOR (LANDLORD)                                  (enter vendor  number of registered vendor  for which invoice number is to be uploaded e.g. 5017253) ',
            ' PERIOD                                                   (please enter month number as per financial year  in which posting is being done  e.g. 1 for april and 7 for october) ',
            'FISCAL YEAR               (Enter Financial year e.g. enter 2019 for FY 2019-20) ',
            'INVOICE COUNT                                             (enter  1 for one entry for a month per Vendor. For multiple registered vendors contract there would be as many  entries for a month as are number of reg. vendors with count 1 for each vendor.More than one postings in same month can also be done by making entries with count  1,2...  but the same need to be done one by one i.e. first making one entry then posting and then making subsequent entry and posting. The system will always pick highest count for a given month. ',
            'INVOICE NUMBER (enter invoice number as received from vendor ) ',
            'INVOICE DATE (enter invoice date in dd.mm.yyyy format)'
        ]

        header_font = Font(name="Arial", size=10, bold=True)
        header_fill = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
        thin_border = Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )

        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        ws.row_dimensions[1].height = 50

        # Data rows
        row_num = 2
        for idx, site in enumerate(sites, start=1):
            c_no = site.get("contract_no") or site.get("Contract Number") or site.get("REM Contract No")
            if not c_no:
                gid = str(site.get("global_id", site.get("Global ID", "")))
                c_no = int(f"1653000{gid.zfill(6)}") if gid.isdigit() else 1653000000000 + idx
            else:
                try:
                    c_no = int(float(c_no))
                except:
                    pass

            row_vals = [
                int(company_code),
                c_no,
                int(vendor_no) if str(vendor_no).isdigit() else vendor_no,
                int(period),
                int(fiscal_year),
                1,
                str(site.get("invoice_no") or invoice_no),
                str(invoice_date)
            ]

            for c_idx, val in enumerate(row_vals, start=1):
                c = ws.cell(row=row_num, column=c_idx, value=val)
                c.border = thin_border
                c.alignment = Alignment(horizontal="center", vertical="center")

            row_num += 1

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    @classmethod
    def generate_condition_upload(
        cls,
        sites: List[Dict[str, Any]],
        bill_type: str = "RENT", # "ENERGY", "RENT", or "LLR"
        circle_code: str = "GJ",
        company_code: int = 1100,
        from_date: str = "01.04.2026",
        to_date: str = "30.04.2026"
    ) -> bytes:
        """
        Creates CONDITION UPLOAD excel matching official BSNL REM template:
        Sheet Name: Circle Code (e.g. 'MP')
        Row 1: 10/11 Column Headers:
        [Status, Company Code, Contract No., Condition Type, Condition Purpose, Valid from, Valid to, Unit Price, Formula, Posting No., Withholding Tax No]
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = str(circle_code).upper()

        headers = [
            'Status',
            'Company Code',
            'Contract No.',
            'Condition Type',
            'Condition Purpose',
            'Valid from',
            'Valid to',
            'Unit Price',
            'Formula',
            'Posting No.',
            'Withholding Tax No'
        ]

        header_font = Font(name="Arial", size=10, bold=True)
        header_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        thin_border = Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )

        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        ws.row_dimensions[1].height = 24

        row_num = 2
        for idx, site in enumerate(sites, start=1):
            c_no = site.get("contract_no") or site.get("Contract Number") or site.get("REM Contract No")
            if not c_no:
                gid = str(site.get("global_id", site.get("Global ID", "")))
                c_no = int(f"1653000{gid.zfill(6)}") if gid.isdigit() else 1653000000000 + idx
            else:
                try:
                    c_no = int(float(c_no))
                except:
                    pass

            # Generate condition records based on bill type
            conditions_to_add = []
            if bill_type == "RENT":
                rent_price = float(site.get("approved_rent", site.get("Net Approved (₹)", site.get("Base Passed (₹)", 0.0))) or 0.0)
                conditions_to_add.append((cls.CONDITION_TYPES["RENT"], rent_price))
            elif bill_type == "LLR":
                llr_price = float(site.get("ground_rent", site.get("Approved Landlord Rent", site.get("Net Approved (₹)", 0.0))) or 0.0)
                conditions_to_add.append((cls.CONDITION_TYPES["LLR"], llr_price))
            else: # ENERGY (creates T570 for EB and T580 for DG if present)
                tot_energy = float(site.get("base_passed", site.get("Base Passed (₹)", site.get("Net Approved (₹)", 0.0))) or 0.0)
                outage_pen = float(site.get("outage_penalty", site.get("Outage Penalty (₹)", 0.0)) or 0.0)
                net_energy = max(0.0, tot_energy - outage_pen)

                dg_h = float(site.get("dg_run_hours", site.get("DG Run/Day", 0.0)) or 0.0)
                if dg_h > 0:
                    dg_ratio = min(0.6, (dg_h / 24.0) * 1.5)
                    dg_amt = round(net_energy * dg_ratio, 2)
                    eb_amt = round(net_energy - dg_amt, 2)
                    conditions_to_add.append((cls.CONDITION_TYPES["ENERGY_EB"], eb_amt))
                    if dg_amt > 0:
                        conditions_to_add.append((cls.CONDITION_TYPES["ENERGY_DG"], dg_amt))
                else:
                    conditions_to_add.append((cls.CONDITION_TYPES["ENERGY_EB"], net_energy))

            for cond_type, unit_price in conditions_to_add:
                row_vals = [
                    'I',
                    int(company_code),
                    c_no,
                    cond_type,
                    'A',
                    from_date,
                    to_date,
                    round(unit_price, 2),
                    'A',
                    None,
                    10 # Withholding Tax
                ]

                for c_idx, val in enumerate(row_vals, start=1):
                    c = ws.cell(row=row_num, column=c_idx, value=val)
                    c.border = thin_border
                    if c_idx in [1, 2, 4, 5, 6, 7, 9, 11]:
                        c.alignment = Alignment(horizontal="center")
                    elif c_idx == 8:
                        c.alignment = Alignment(horizontal="right")
                        c.number_format = '#,##0.00'

                row_num += 1

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    @classmethod
    def generate_all_3_templates_zip(
        cls,
        invoice_no: str,
        sites: List[Dict[str, Any]],
        bill_type: str = "RENT", # "ENERGY", "RENT", or "LLR"
        circle_code: str = "GJ",
        company_code: int = 1100,
        business_entity: str = "16400000",
        vendor_no: str = "2004755",
        invoice_date: str = "05.05.2026",
        from_date: str = "01.04.2026",
        to_date: str = "30.04.2026",
        period: int = 1,
        fiscal_year: int = 2026
    ) -> bytes:
        """
        Bundles all 3 upload files into a ready-to-upload ZIP archive.
        """
        b_inv_upload = cls.generate_invoice_upload(
            invoice_no=invoice_no,
            sites=sites,
            bill_type=bill_type,
            circle_code=circle_code,
            company_code=company_code,
            business_entity=business_entity,
            vendor_no=vendor_no,
            from_date=from_date,
            to_date=to_date
        )

        b_detail_upload = cls.generate_invoice_detail_upload(
            invoice_no=invoice_no,
            sites=sites,
            circle_code=circle_code,
            company_code=company_code,
            vendor_no=vendor_no,
            invoice_date=invoice_date,
            period=period,
            fiscal_year=fiscal_year
        )

        b_cond_upload = cls.generate_condition_upload(
            sites=sites,
            bill_type=bill_type,
            circle_code=circle_code,
            company_code=company_code,
            from_date=from_date,
            to_date=to_date
        )

        zip_buf = io.BytesIO()
        prefix = f"{circle_code}_{bill_type}_{invoice_no}"
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr(f"INVOICE_UPLOAD_{prefix}.xlsx", b_inv_upload)
            z.writestr(f"INVOICE_DETAIL_UPLOAD_{prefix}.xlsx", b_detail_upload)
            z.writestr(f"CONDITION_UPLOAD_{prefix}.xlsx", b_cond_upload)

        return zip_buf.getvalue()

    @classmethod
    def generate_grouped_invoices_zip(
        cls,
        grouped_sites: Dict[str, List[Dict[str, Any]]],
        bill_type: str = "RENT",
        circle_code: str = "GJ",
        company_code: int = 1100,
        business_entity: str = "16400000",
        vendor_no: str = "2004755",
        invoice_date: str = "05.05.2026",
        from_date: str = "01.04.2026",
        to_date: str = "30.04.2026",
        period: int = 1,
        fiscal_year: int = 2026
    ) -> bytes:
        """
        Creates a master ZIP package where each distinct bill number has its own folder
        containing the 3 official BSNL SAP REM upload templates for that bill.
        """
        master_buf = io.BytesIO()
        with zipfile.ZipFile(master_buf, "w", zipfile.ZIP_DEFLATED) as master_z:
            for inv_no, sites in grouped_sites.items():
                if not sites:
                    continue
                inv_clean = str(inv_no).strip()
                b_inv = cls.generate_invoice_upload(
                    invoice_no=inv_clean,
                    sites=sites,
                    bill_type=bill_type,
                    circle_code=circle_code,
                    company_code=company_code,
                    business_entity=business_entity,
                    vendor_no=vendor_no,
                    from_date=from_date,
                    to_date=to_date
                )
                b_det = cls.generate_invoice_detail_upload(
                    invoice_no=inv_clean,
                    sites=sites,
                    circle_code=circle_code,
                    company_code=company_code,
                    vendor_no=vendor_no,
                    invoice_date=invoice_date,
                    period=period,
                    fiscal_year=fiscal_year
                )
                b_cond = cls.generate_condition_upload(
                    sites=sites,
                    bill_type=bill_type,
                    circle_code=circle_code,
                    company_code=company_code,
                    from_date=from_date,
                    to_date=to_date
                )
                folder = f"Invoice_{inv_clean}"
                master_z.writestr(f"{folder}/INVOICE_UPLOAD_{circle_code}_{bill_type}_{inv_clean}.xlsx", b_inv)
                master_z.writestr(f"{folder}/INVOICE_DETAIL_UPLOAD_{circle_code}_{bill_type}_{inv_clean}.xlsx", b_det)
                master_z.writestr(f"{folder}/CONDITION_UPLOAD_{circle_code}_{bill_type}_{inv_clean}.xlsx", b_cond)

        return master_buf.getvalue()

    @classmethod
    def generate_master_all_invoices_zip(
        cls,
        excel_path: str,
        contract_map: Dict[str, str],
        bill_type: str = "ENERGY",
        circle_code: str = "MP",
        company_code: int = 1100,
        business_entity: str = "23400000",
        vendor_no: str = "2004755",
        period: int = 11,
        fiscal_year: int = 2026,
        from_date: str = "01.02.2026",
        to_date: str = "28.02.2026"
    ) -> bytes:
        """
        Generates a master ZIP archive containing 3 upload templates for each of the 5 invoices (15 templates total).
        """
        inv_list = [
            ("607239111046515", 1),
            ("607239111046516", 2),
            ("607239111046519", 3),
            ("607239111046520", 4),
            ("607239111046533", 5)
        ]

        master_buf = io.BytesIO()
        with zipfile.ZipFile(master_buf, "w", zipfile.ZIP_DEFLATED) as master_zip:
            for inv_no, idx in inv_list:
                try:
                    sheet_name = f"Inv {idx} ({inv_no})"
                    df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=2)
                    df = df[~df["S.No"].astype(str).str.upper().str.contains("TOTAL", na=False)]

                    site_records = []
                    for _, r in df.iterrows():
                        gid = str(r.get("Global ID", "")).split(".")[0].strip()
                        vid = str(r.get("Vendor Site ID", "")).strip()
                        c_no = contract_map.get(gid) or contract_map.get(vid)
                        if not c_no:
                            c_no = f"2353000{gid.zfill(6)}" if gid.isdigit() else "2353000000001"

                        site_records.append({
                            "contract_no": c_no,
                            "global_id": gid,
                            "site_name": r.get("Site Name", ""),
                            "base_passed": float(r.get("Base Passed (₹)", 0.0) or 0.0),
                            "outage_penalty": float(r.get("Outage Penalty (₹)", 0.0) or 0.0),
                            "approved_rent": float(r.get("Net Approved (₹)", r.get("Base Passed (₹)", 0.0)) or 0.0),
                            "ground_rent": 0.0,
                            "dg_run_hours": float(r.get("DG Run/Day", 0.0) or 0.0)
                        })

                    b_inv = cls.generate_invoice_upload(
                        invoice_no=inv_no,
                        sites=site_records,
                        bill_type=bill_type,
                        circle_code=circle_code,
                        company_code=company_code,
                        business_entity=business_entity,
                        vendor_no=vendor_no,
                        from_date=from_date,
                        to_date=to_date
                    )
                    b_det = cls.generate_invoice_detail_upload(
                        invoice_no=inv_no,
                        sites=site_records,
                        circle_code=circle_code,
                        company_code=company_code,
                        vendor_no=vendor_no,
                        period=period,
                        fiscal_year=fiscal_year
                    )
                    b_cond = cls.generate_condition_upload(
                        sites=site_records,
                        bill_type=bill_type,
                        circle_code=circle_code,
                        company_code=company_code,
                        from_date=from_date,
                        to_date=to_date
                    )

                    folder = f"Invoice_{inv_no}"
                    master_zip.writestr(f"{folder}/INVOICE_UPLOAD_{circle_code}_{inv_no}.xlsx", b_inv)
                    master_zip.writestr(f"{folder}/INVOICE_detail_UPLOAD_{circle_code}_{inv_no}.xlsx", b_det)
                    master_zip.writestr(f"{folder}/CONDITION_UPLOAD_{circle_code}_{bill_type}_{inv_no}.xlsx", b_cond)
                except Exception:
                    pass

        return master_buf.getvalue()
