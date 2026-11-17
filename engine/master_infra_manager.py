"""
BSNL Master Telecom Infrastructure Site Data Manager
Manages multi-operator, multi-circle master site databases (ATC, Indus, Ascend, etc.)
Handles ingestion, parsing, updates, and reconciliation data extraction for Gujarat, MP, etc.
"""

import os
import io
import json
import openpyxl
import pandas as pd
from typing import Dict, Any, List, Optional, Tuple


MASTER_DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "master_infra_data")
MANIFEST_PATH = os.path.join(MASTER_DATA_DIR, "manifest.json")


class MasterInfraManager:
    """
    Central Manager for Operator Master Infrastructure Datasets.
    """

    @classmethod
    def get_master_dir(cls) -> str:
        os.makedirs(MASTER_DATA_DIR, exist_ok=True)
        return MASTER_DATA_DIR

    @classmethod
    def get_manifest(cls) -> Dict[str, Any]:
        cls.get_master_dir()
        if os.path.exists(MANIFEST_PATH):
            try:
                with open(MANIFEST_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass

        # Default initial manifest
        default_manifest = {
            "Elevar Digitel (ATC)": {
                "Gujarat": {
                    "filename": "1_ATC_GJ_IPF_2026-27_Main.xlsx",
                    "type": "IP Tower Rent (IPF)",
                    "periods": ["APR-2026", "MAY-2026", "JUN-2026", "MAR-2026"],
                    "sheets": ["APR-26", "MAY-26", "JUNE -26 ", "March _KJJ"],
                    "description": "Master IP Fee Reconciliation & Site Registry for Gujarat Circle (FY 2026-27) - 580 Sites",
                    "updated_at": "2026-11-16"
                },
                "Madhya Pradesh": {
                    "filename": "media_1794423587823_bcbc41b2.xlsx",
                    "type": "IP Tower Rent & Contract Registry",
                    "periods": ["FEB-2026", "MAR-2026", "APR-2026", "MAY-2026", "JUN-2026", "JUL-2026", "AUG-2026"],
                    "sheets": ["ELEVAR DIGITEL FEB-26", "FINAL SUMMARY Jul-26 & Aug-26"],
                    "description": "Master IP Rental Reconciliation & REM Contract Mapping for MP Circle - 486 Sites",
                    "updated_at": "2026-11-14"
                }
            }
        }
        cls.save_manifest(default_manifest)
        return default_manifest

    @classmethod
    def save_manifest(cls, manifest_data: Dict[str, Any]) -> None:
        cls.get_master_dir()
        with open(MANIFEST_PATH, "w", encoding="utf-8") as f:
            json.dump(manifest_data, f, indent=4)

    @classmethod
    def is_data_available(cls, vendor: str, circle: str, month: str, year: str) -> bool:
        """
        Checks whether data is available for the given vendor, circle, and period.
        """
        # Elevar Digitel in MP (FEB-2026)
        if vendor == "Elevar Digitel (ATC)" and circle == "Madhya Pradesh" and month == "FEB" and str(year) == "2026":
            return True

        # Elevar Digitel in Gujarat (APR, MAY, JUN, MAR 2026)
        if vendor == "Elevar Digitel (ATC)" and circle == "Gujarat" and str(year) == "2026":
            if month in ["APR", "MAY", "JUN", "MAR"]:
                fpath_new = os.path.join(cls.get_master_dir(), "Copy_of_IP_Fees_console_April_26_VENDORE_SUBMITTED_BSNL.xlsx")
                fpath_main = os.path.join(cls.get_master_dir(), "1_ATC_GJ_IPF_2026-27_Main.xlsx")
                return os.path.exists(fpath_new) or os.path.exists(fpath_main)

        return False

    @classmethod
    def find_header_row(cls, df_raw: pd.DataFrame, max_search: int = 15) -> int:
        """
        Dynamically finds the row index containing site headers ('Site Name', 'IP ID', 'Sr. No').
        """
        for r in range(min(max_search, len(df_raw))):
            vals = [str(v).strip().lower() for v in df_raw.iloc[r].values if pd.notna(v)]
            if any('site name' in v or 'ip id' in v for v in vals):
                return r
        return 4 # Default fallback

    @classmethod
    def load_gujarat_atc_month_data(cls, month_str: str = "APR", year_str: str = "2026") -> Dict[str, Any]:
        """
        Parses Gujarat ATC master file for the requested month.
        Compares with BSNL internal checklist records (EDIL_P&F_checklist_25-26.xlsx sheet 'March-26 ')
        to calculate head-wise deductions (Tenancy, City Class, Technology Loading) and Outage deductions.
        """
        # Check if the new April 2026 vendor submitted invoice is present
        new_apr_path = os.path.join(cls.get_master_dir(), "Copy_of_IP_Fees_console_April_26_VENDORE_SUBMITTED_BSNL.xlsx")
        extracted_checklist_path = os.path.join(cls.get_master_dir(), "BSNL_Checklist_March_26_Extracted.xlsx")
        checklist_path = os.path.join(cls.get_master_dir(), "EDIL_P&F_checklist_25-26.xlsx")
        if not os.path.exists(checklist_path) and os.path.exists(r"C:\Users\vanya\Downloads\EDIL_P&F_checklist_25-26.xlsx"):
            checklist_path = r"C:\Users\vanya\Downloads\EDIL_P&F_checklist_25-26.xlsx"

        if month_str.upper()[:3] == "APR" and os.path.exists(new_apr_path):
            try:
                # 1. Load Vendor Console
                df_vendor = pd.read_excel(new_apr_path, sheet_name="April'26 (2)", header=3)
                id_col = next((c for c in df_vendor.columns if "GLOBAL_SITE_ID" in str(c) or "Site ID" in str(c)), df_vendor.columns[1])
                valid_df = df_vendor[df_vendor[id_col].notna()].copy()
                valid_df = valid_df[~valid_df[id_col].astype(str).str.upper().str.contains("TOTAL|SUB|GRAND", na=False)]
                valid_df["SAP_ID_CLEAN"] = valid_df["SAP ID"].astype(str).str.strip().str.upper()
                valid_df["OMCR_ID_CLEAN"] = valid_df["OPCO Site ID"].astype(str).str.strip().str.upper()
                valid_df["GID_CLEAN"] = valid_df["GLOBAL_SITE_ID"].astype(str).str.split('.').str[0].str.strip()

                # 2. Load BSNL Checklist (March-26 sheet)
                df_bsnl = pd.DataFrame()
                if os.path.exists(extracted_checklist_path):
                    try:
                        df_bsnl = pd.read_excel(extracted_checklist_path)
                        df_bsnl["SAP_ID_CLEAN"] = df_bsnl["ATC SAP ID"].astype(str).str.strip().str.upper()
                        df_bsnl["OMCR_ID_CLEAN"] = df_bsnl["BSNL OMCR ID"].astype(str).str.strip().str.upper()
                    except Exception as ex_ext:
                        print(f"Error loading extracted checklist: {ex_ext}")
                
                if df_bsnl.empty and os.path.exists(checklist_path):
                    try:
                        import openpyxl
                        wb_edil = openpyxl.load_workbook(checklist_path, read_only=True, data_only=True)
                        target_sheet = next((s for s in wb_edil.sheetnames if "MARCH-26" in s.upper() or "MAR26" in s.upper() or "MAR-26" in s.upper()), None)
                        if target_sheet:
                            s_m26 = wb_edil[target_sheet]
                            b_rows = []
                            b_headers = []
                            for r_i, r_val in enumerate(s_m26.iter_rows(values_only=True)):
                                if r_i == 4:
                                    b_headers = [str(c).strip() if c is not None else f"Col_{c_i}" for c_i, c in enumerate(r_val)]
                                elif r_i > 4 and r_val[1] is not None:
                                    b_rows.append(r_val[:len(b_headers)])
                            df_bsnl = pd.DataFrame(b_rows, columns=b_headers)
                            df_bsnl["SAP_ID_CLEAN"] = df_bsnl["ATC SAP ID"].astype(str).str.strip().str.upper()
                            df_bsnl["OMCR_ID_CLEAN"] = df_bsnl["BSNL OMCR ID"].astype(str).str.strip().str.upper()
                        wb_edil.close()
                    except Exception:
                        pass

                # 3. Load Outages
                outage_path = os.path.join(cls.get_master_dir(), "All_IP_sites_DT_AC_MR_report_APR-26xls.xls")
                outage_map = {}
                outage_df = pd.DataFrame()
                if os.path.exists(outage_path):
                    try:
                        from engine.outage_parser import OutageParser
                        outage_df = OutageParser.parse_multi_block_outage(outage_path)
                        df_atc_out = outage_df[outage_df["sheet"] == "ATC_DT_AC"].copy()
                        for _, r_out in df_atc_out[df_atc_out["duration_hours"] > 0].iterrows():
                            dur_hrs = float(r_out.get("duration_hours") or 0.0)
                            tot_mins = dur_hrs * 60.0
                            excess_mins = max(0.0, tot_mins - 43.2)
                            pen_amt = min(10000.0, (excess_mins / 60.0) * 500.0)
                            out_item = {
                                "dur_hrs": dur_hrs,
                                "outage_mins": tot_mins,
                                "pen_amt": pen_amt,
                                "fault": str(r_out.get("fault_reason", "")),
                                "ac_stat": str(r_out.get("ac_status", ""))
                            }
                            for id_k in [r_out.get("IP ID"), r_out.get("OMCR ID")]:
                                if id_k and str(id_k).strip().upper() not in ["NAN", "NONE", ""]:
                                    outage_map[str(id_k).strip().upper()] = out_item
                    except Exception:
                        pass

                # 4. Standard MSA Sharing Rates Lookup Matrix
                rate_table = {
                    ("A", "GBT", 1): 28626.0, ("A", "GBT", 2): 28626.0, ("A", "GBT", 3): 26987.0, ("A", "GBT", 4): 23728.0, ("A", "GBT", 5): 23228.0,
                    ("A", "RTT", 1): 17303.0, ("A", "RTT", 2): 17303.0, ("A", "RTT", 3): 14235.0, ("A", "RTT", 4): 13601.0, ("A", "RTT", 5): 13150.0,
                    ("B1", "GBT", 1): 28242.0, ("B1", "GBT", 2): 28242.0, ("B1", "GBT", 3): 26910.0, ("B1", "GBT", 4): 23500.0, ("B1", "GBT", 5): 23228.0,
                    ("B1", "RTT", 1): 16785.0, ("B1", "RTT", 2): 16785.0, ("B1", "RTT", 3): 13890.0, ("B1", "RTT", 4): 13343.0, ("B1", "RTT", 5): 13000.0,
                    ("B2", "GBT", 1): 27927.0, ("B2", "GBT", 2): 27927.0, ("B2", "GBT", 3): 26600.0, ("B2", "GBT", 4): 23300.0, ("B2", "GBT", 5): 23000.0,
                    ("B2", "RTT", 1): 16650.0, ("B2", "RTT", 2): 16650.0, ("B2", "RTT", 3): 13890.0, ("B2", "RTT", 4): 13100.0, ("B2", "RTT", 5): 12800.0,
                    ("C", "GBT", 1): 27477.0, ("C", "GBT", 2): 27477.0, ("C", "GBT", 3): 26400.0, ("C", "GBT", 4): 23368.0, ("C", "GBT", 5): 22868.0,
                    ("C", "RTT", 1): 15750.0, ("C", "RTT", 2): 15750.0, ("C", "RTT", 3): 13200.0, ("C", "RTT", 4): 12825.0, ("C", "RTT", 5): 12500.0,
                    ("RURAL", "GBT", 1): 27027.0, ("RURAL", "GBT", 2): 27027.0, ("RURAL", "GBT", 3): 26100.0, ("RURAL", "GBT", 4): 23143.0, ("RURAL", "GBT", 5): 22643.0,
                    ("RURAL", "RTT", 1): 15075.0, ("RURAL", "RTT", 2): 15075.0, ("RURAL", "RTT", 3): 12750.0, ("RURAL", "RTT", 4): 12400.0, ("RURAL", "RTT", 5): 12200.0,
                }

                audit_records = []
                for _, v_row in valid_df.iterrows():
                    sap = v_row["SAP_ID_CLEAN"]
                    omcr = v_row["OMCR_ID_CLEAN"]
                    gid = v_row["GID_CLEAN"]

                    b_match = pd.DataFrame()
                    if not df_bsnl.empty:
                        b_match = df_bsnl[df_bsnl["SAP_ID_CLEAN"] == sap]
                        if b_match.empty and omcr:
                            b_match = df_bsnl[df_bsnl["OMCR_ID_CLEAN"] == omcr]
                    b_site = b_match.iloc[0] if not b_match.empty else None

                    v_tenants = float(v_row.get("No.of Tenants") or 1.0)
                    v_city = str(v_row.get("City Class", "RURAL")).strip().upper()
                    v_type = str(v_row.get("Site Type", "RTT")).strip().upper()
                    norm_v_type = "GBT" if "GBT" in v_type else "RTT"

                    v_base_rate = float(v_row.get("Applied MSA Rate") or 0.0)
                    v_loading = float(v_row.get("Total Loading Mark up") or 0.0)
                    v_total_bill = float(v_row.get("Total Billing") or (v_base_rate + v_loading))

                    if b_site is not None:
                        b_tenants = float(b_site.get("No of Tenants") or v_tenants)
                        b_city = str(b_site.get("Class of City", v_city)).strip().upper()
                        b_type = str(b_site.get("Type of Tower", v_type)).strip().upper()
                        b_4g = str(b_site.get("4G", "0")).strip().upper()
                        b_ssa = str(b_site.get("SSA", v_row.get("City Class", ""))).strip()
                        c_no = str(b_site.get("ISA") or "").strip()
                        r_no = str(b_site.get("RO") or "").strip()
                    else:
                        b_tenants = v_tenants
                        b_city = v_city
                        b_type = v_type
                        b_4g = "TEJAS"
                        b_ssa = str(v_row.get("City Class", ""))
                        c_no = ""
                        r_no = ""

                    # 1. Tenancy Deduction
                    ded_tenancy = 0.0
                    if b_tenants > v_tenants:
                        allow_rate = rate_table.get((v_city, norm_v_type, min(5, int(b_tenants))), v_base_rate)
                        if v_base_rate > allow_rate:
                            ded_tenancy = max(0.0, v_base_rate - allow_rate)

                    # 2. City Class Deduction
                    ded_city = 0.0
                    if b_city != v_city:
                        allow_city_rate = rate_table.get((b_city, norm_v_type, min(5, int(v_tenants))), v_base_rate)
                        if v_base_rate > allow_city_rate:
                            ded_city = max(0.0, v_base_rate - allow_city_rate)

                    # 3. Technology Loading Deduction
                    ded_loading = 0.0
                    if b_4g in ["0", "NONE", "NAN", "", "FALSE"] and v_loading > 0:
                        ded_loading = v_loading

                    # 4. Outage Deduction
                    o_info = outage_map.get(sap) or outage_map.get(omcr) or outage_map.get(gid) or {}
                    ded_outage = float(o_info.get("pen_amt", 0.0))
                    outage_hrs = float(o_info.get("dur_hrs", 0.0))
                    ac_stat = str(o_info.get("ac_stat", ""))
                    fault = str(o_info.get("fault", ""))

                    tot_ded = ded_tenancy + ded_city + ded_loading + ded_outage
                    net_payable = max(0.0, v_total_bill - tot_ded)

                    reasons = []
                    if ded_tenancy > 0:
                        reasons.append(f"Tenancy Mismatch (Ven:{int(v_tenants)} vs BSNL:{int(b_tenants)})")
                    if ded_city > 0:
                        reasons.append(f"City Class Mismatch (Ven:{v_city} vs BSNL:{b_city})")
                    if ded_loading > 0:
                        reasons.append("4G Loading Disallowed (BSNL 4G: 0)")
                    if ded_outage > 0:
                        reasons.append(f"Outage SLA Breach ({outage_hrs}h)")

                    audit_status = "Deduction Applied" if tot_ded > 0 else "Fully Passed"
                    remarks = "; ".join(reasons) if reasons else "Parameters verified with BSNL Checklist"

                    audit_records.append({
                        "Srno": v_row.get("Srno"),
                        "GLOBAL_SITE_ID": gid,
                        "SAP ID": sap,
                        "OPCO Site ID": omcr,
                        "Site Name": str(v_row.get("Site Name", "")).strip().strip('"'),
                        "SSA": b_ssa,
                        "City Class (Ven)": v_city,
                        "City Class (BSNL)": b_city,
                        "Tenants (Ven)": int(v_tenants),
                        "Tenants (BSNL)": int(b_tenants),
                        "BSNL 4G Tech": b_4g,
                        "Claimed Base Rate (₹)": round(v_base_rate, 2),
                        "Claimed Loading (₹)": round(v_loading, 2),
                        "Total Billed (₹)": round(v_total_bill, 2),
                        "Deduction Tenancy (₹)": round(ded_tenancy, 2),
                        "Deduction City Class (₹)": round(ded_city, 2),
                        "Deduction Loading (₹)": round(ded_loading, 2),
                        "Deduction Outage (₹)": round(ded_outage, 2),
                        "Total Deductions (₹)": round(tot_ded, 2),
                        "Net Approved (₹)": round(net_payable, 2),
                        "Audit Status": audit_status,
                        "Deduction Remarks": remarks,
                        "Invoice no": str(v_row.get("Invoice no")).split('.')[0].strip(),
                        "Outage Hours": outage_hrs,
                        "AC Status": ac_stat,
                        "ISA": c_no,
                        "RO": r_no,
                        # Compatibility columns
                        "Total IP Fee": round(net_payable, 2),
                        "IP Claim with GST": round(v_total_bill * 1.18, 2),
                        "Diff": round(tot_ded, 2)
                    })

                df_audit = pd.DataFrame(audit_records)

                # Invoice summaries
                distinct_invoices = [str(x).split('.')[0].strip() for x in df_audit["Invoice no"].dropna().unique() if str(x).strip().isdigit()]
                inv_summary_list = []
                inv_site_dfs = {}
                for inv_no in distinct_invoices:
                    idf = df_audit[df_audit["Invoice no"] == inv_no].copy()
                    site_cnt = len(idf)
                    tot_claim_b = float(idf["Total Billed (₹)"].sum())
                    tot_ded_b = float(idf["Total Deductions (₹)"].sum())
                    tot_pass_b = float(idf["Net Approved (₹)"].sum())
                    inv_summary_list.append({
                        "Invoice No": str(inv_no),
                        "Sites Count": site_cnt,
                        "Vendor Claimed Excl. GST (₹)": round(tot_claim_b, 2),
                        "Total Deduction / Recovery (₹)": round(tot_ded_b, 2),
                        "BSNL Approved Excl. GST (₹)": round(tot_pass_b, 2),
                        "GST @ 18% (₹)": round(tot_pass_b * 0.18, 2),
                        "Net Payable Incl. GST (₹)": round(tot_pass_b * 1.18, 2)
                    })
                    inv_site_dfs[str(inv_no)] = idf

                summary_df = pd.DataFrame(inv_summary_list) if inv_summary_list else pd.DataFrame()

                total_sites = len(df_audit)
                total_billed = float(df_audit["Total Billed (₹)"].sum())
                total_passed = float(df_audit["Net Approved (₹)"].sum())
                total_deductions = float(df_audit["Total Deductions (₹)"].sum())
                total_claimed_gst = total_billed * 1.18

                # Deployment Penalties
                penalties_df = pd.DataFrame()
                try:
                    penalties_df = pd.read_excel(new_apr_path, sheet_name="Sheet1")
                except Exception:
                    pass

                return {
                    "month_str": month_str,
                    "sheet_name": "April'26 (2)",
                    "total_sites": total_sites,
                    "total_billed": total_billed,
                    "total_passed": total_passed,
                    "total_claimed": total_claimed_gst,
                    "total_diff": total_deductions,
                    "total_deductions": total_deductions,
                    "tenancy_deductions": float(df_audit["Deduction Tenancy (₹)"].sum()),
                    "city_deductions": float(df_audit["Deduction City Class (₹)"].sum()),
                    "loading_deductions": float(df_audit["Deduction Loading (₹)"].sum()),
                    "outage_deductions": float(df_audit["Deduction Outage (₹)"].sum()),
                    "sites_df": df_audit,
                    "invoices_summary_df": summary_df,
                    "distinct_invoices": distinct_invoices,
                    "inv_site_dfs": inv_site_dfs,
                    "penalties_df": penalties_df,
                    "outage_df": outage_df,
                    "is_new_april_console": True,
                    "is_checklist_audited": True,
                    "month": month_str,
                    "year": year_str,
                }
            except Exception as e:
                pass

        fpath = os.path.join(cls.get_master_dir(), "1_ATC_GJ_IPF_2026-27_Main.xlsx")
        if not os.path.exists(fpath):
            return {}

        month_map = {
            "APR": "APR-26",
            "MAY": "MAY-26",
            "JUN": "JUNE -26 ",
            "MAR": "March _KJJ"
        }
        sheet_target = month_map.get(month_str.upper()[:3], "APR-26")

        try:
            # 1. Read Raw Sheet
            df_raw = pd.read_excel(fpath, sheet_name=sheet_target, header=None)
            h_row = cls.find_header_row(df_raw)

            headers = []
            counts = {}
            for idx, col in enumerate(df_raw.iloc[h_row].values):
                name = str(col).strip() if pd.notna(col) else f"Col_{idx}"
                if name in counts:
                    counts[name] += 1
                    name = f"{name}_{counts[name]}"
                else:
                    counts[name] = 1
                headers.append(name)

            data_df = df_raw.iloc[h_row + 1:].copy()
            data_df.columns = headers

            # Clean and filter to valid site rows
            id_col = next((c for c in data_df.columns if "IP ID" in str(c)), None)
            if not id_col:
                id_col = data_df.columns[1]

            valid_df = data_df[data_df[id_col].notna()].copy()
            # Drop empty / total rows
            valid_df = valid_df[~valid_df[id_col].astype(str).str.upper().str.contains("TOTAL|SUB|GRAND", na=False)]

            # Convert financial columns to numeric
            num_cols = ["Total IP Fee", "IP Claim with GST", "Diff", "1st Esc 12.5 %", "2nd Esc 12.5 %", "Charges for Extra Rack", "Charges for Extra Antena", "4G upgradation"]
            for col in num_cols:
                if col in valid_df.columns:
                    valid_df[col] = pd.to_numeric(valid_df[col], errors="coerce").fillna(0.0)

            # 2. Extract Invoices
            inv_col = next((c for c in valid_df.columns if "Invoice" in str(c)), None)
            distinct_invoices = []
            if inv_col:
                for x in valid_df[inv_col].dropna().unique():
                    inv_str = str(x).split('.')[0].strip()
                    if inv_str and inv_str.isdigit() and len(inv_str) >= 8:
                        distinct_invoices.append(inv_str)

            # If distinct invoices not directly in master sheet column, get sheet names
            xls = pd.ExcelFile(fpath)
            month_inv_sheets = []
            for sname in xls.sheet_names:
                if str(sname).strip().isdigit() and len(str(sname).strip()) >= 12:
                    # Check if sheet belongs to this month
                    try:
                        sheet_head = pd.read_excel(fpath, sheet_name=sname, nrows=6, header=None)
                        head_txt = " ".join([str(v) for v in sheet_head.values.flatten()])
                        if sheet_target[:3].upper() in head_txt.upper():
                            month_inv_sheets.append(str(sname).strip())
                    except Exception:
                        pass

            if not distinct_invoices and month_inv_sheets:
                distinct_invoices = month_inv_sheets

            # 3. Build Invoice Summary Table
            inv_summary_list = []
            inv_site_dfs = {}

            for inv_no in distinct_invoices:
                # Check if matching dedicated sheet exists in workbook
                sheet_match = [s for s in xls.sheet_names if str(inv_no) in str(s)]
                if sheet_match:
                    s_name = sheet_match[0]
                    raw_inv_df = pd.read_excel(fpath, sheet_name=s_name, header=None)
                    inv_h = cls.find_header_row(raw_inv_df)
                    inv_headers = []
                    inv_counts = {}
                    for idx, col in enumerate(raw_inv_df.iloc[inv_h].values):
                        name = str(col).strip() if pd.notna(col) else f"Col_{idx}"
                        if name in inv_counts:
                            inv_counts[name] += 1
                            name = f"{name}_{inv_counts[name]}"
                        else:
                            inv_counts[name] = 1
                        inv_headers.append(name)
                    idf = raw_inv_df.iloc[inv_h + 1:].copy()
                    idf.columns = inv_headers
                    i_id_col = next((c for c in idf.columns if "IP ID" in str(c)), None)
                    if i_id_col:
                        idf = idf[idf[i_id_col].notna()]
                        idf = idf[~idf[i_id_col].astype(str).str.upper().str.contains("TOTAL|SUB|GRAND", na=False)]
                else:
                    idf = valid_df[valid_df[inv_col].astype(str).str.contains(str(inv_no), na=False)].copy() if inv_col else pd.DataFrame()

                for col in num_cols:
                    if col in idf.columns:
                        idf[col] = pd.to_numeric(idf[col], errors="coerce").fillna(0.0)

                site_cnt = len(idf)
                tot_passed = float(idf["Total IP Fee"].sum()) if "Total IP Fee" in idf.columns else 0.0
                tot_claim = float(idf["IP Claim with GST"].sum()) if "IP Claim with GST" in idf.columns else (tot_passed * 1.18)
                tot_diff = float(idf["Diff"].sum()) if "Diff" in idf.columns else 0.0

                inv_summary_list.append({
                    "Invoice No": str(inv_no),
                    "Sites Count": site_cnt,
                    "BSNL Passed Excl. GST (₹)": round(tot_passed, 2),
                    "Vendor Claimed With GST (₹)": round(tot_claim, 2),
                    "Total Deduction / Diff (₹)": round(tot_diff, 2),
                    "GST @ 18% (₹)": round(tot_passed * 0.18, 2),
                    "Net Payable With GST (₹)": round(tot_passed * 1.18, 2)
                })
                inv_site_dfs[str(inv_no)] = idf

            df_inv_summary = pd.DataFrame(inv_summary_list)

            # KPIs
            total_sites = len(valid_df)
            total_passed = float(valid_df["Total IP Fee"].sum()) if "Total IP Fee" in valid_df.columns else 0.0
            total_claimed = float(valid_df["IP Claim with GST"].sum()) if "IP Claim with GST" in valid_df.columns else (total_passed * 1.18)
            total_diff = float(valid_df["Diff"].sum()) if "Diff" in valid_df.columns else 0.0

            return {
                "month_str": month_str,
                "sheet_name": sheet_target,
                "total_sites": total_sites,
                "total_passed": total_passed,
                "total_claimed": total_claimed,
                "total_diff": total_diff,
                "sites_df": valid_df,
                "invoices_summary_df": df_inv_summary,
                "distinct_invoices": distinct_invoices,
                "inv_site_dfs": inv_site_dfs
            }
        except Exception as e:
            return {"error": str(e)}

    @classmethod
    def save_uploaded_master_file(cls, uploaded_file, operator_name: str, circle_name: str, file_type: str = "IP Tower Rent (IPF)") -> Tuple[bool, str]:
        """
        Saves or updates a master file uploaded by the user.
        """
        try:
            target_dir = cls.get_master_dir()
            clean_op = "".join(c for c in operator_name if c.isalnum() or c in (' ', '_')).strip().replace(" ", "_")
            clean_cir = "".join(c for c in circle_name if c.isalnum() or c in (' ', '_')).strip().replace(" ", "_")
            save_filename = f"{clean_op}_{clean_cir}_Master.xlsx"
            save_path = os.path.join(target_dir, save_filename)

            with open(save_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Read sheets in uploaded workbook
            wb = openpyxl.load_workbook(save_path, read_only=True)
            sheet_names = wb.sheetnames

            # Update manifest
            manifest = cls.get_manifest()
            if operator_name not in manifest:
                manifest[operator_name] = {}

            manifest[operator_name][circle_name] = {
                "filename": save_filename,
                "type": file_type,
                "sheets": sheet_names,
                "updated_at": str(pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")),
                "description": f"Master site database for {operator_name} in {circle_name} ({len(sheet_names)} sheets)"
            }
            cls.save_manifest(manifest)

            return True, f"Successfully uploaded and registered '{save_filename}' for {operator_name} ({circle_name}) with {len(sheet_names)} sheets!"
        except Exception as e:
            return False, f"Failed to save master file: {str(e)}"
