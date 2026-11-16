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

        # Elevar Digitel in Gujarat (APR, MAY, JUN 2026)
        if vendor == "Elevar Digitel (ATC)" and circle == "Gujarat" and str(year) == "2026":
            if month in ["APR", "MAY", "JUN"]:
                fpath = os.path.join(cls.get_master_dir(), "1_ATC_GJ_IPF_2026-27_Main.xlsx")
                return os.path.exists(fpath)

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
        Returns master sites DataFrame, invoice summary DataFrame, and per-invoice site DataFrames.
        """
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
