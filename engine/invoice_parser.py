"""
Invoice and Site Annexure Parser for BSNL IP Vendor Invoices
Supports PDF invoices and Excel Site Annexures (Elevar / ATC / Indus / Ascend)
"""

import re
import os
from typing import Dict, List, Any, Optional
import pandas as pd



class InvoiceParser:
    def __init__(self):
        pass

    @staticmethod
    def parse_pdf_invoice_summary(pdf_path: str) -> List[Dict[str, Any]]:
        """
        Extract high-level Tax Invoice headers from PDF.
        """
        invoices = []
        try:
            import fitz
        except Exception:
            try:
                import pymupdf as fitz
            except Exception:
                fitz = None

        if fitz is None:
            raise ImportError("PyMuPDF is not properly installed. Please run: pip install PyMuPDF==1.24.14")
        doc = fitz.open(pdf_path)

        for page_idx in range(len(doc)):
            page = doc[page_idx]
            text = page.get_text()

            # If scanned, use RapidOCR
            if not text.strip():
                try:
                    from rapidocr_onnxruntime import RapidOCR
                    engine = RapidOCR()
                    pix = page.get_pixmap(dpi=150)
                    img_bytes = pix.tobytes("png")
                    ocr_res, _ = engine(img_bytes)
                    if ocr_res:
                        text = "\n".join([line[1] for line in ocr_res])
                except Exception:
                    continue

            if "TAX INVOICE" in text.upper() or "INVOICE" in text.upper():
                inv_data = InvoiceParser._extract_invoice_fields(text, page_idx + 1)
                if inv_data.get("invoice_no"):
                    invoices.append(inv_data)

        return invoices

    @staticmethod
    def _extract_invoice_fields(text: str, page_num: int) -> Dict[str, Any]:
        """Extract individual fields from invoice text."""
        data = {"page": page_num}

        # Invoice number pattern
        inv_match = re.search(r'(?:Number|Invoice\s*No\.?)\s*[:\s]*([0-9]{10,18})', text, re.IGNORECASE)
        if inv_match:
            data["invoice_no"] = inv_match.group(1).strip()
        else:
            long_nums = re.findall(r'\b60723[0-9]{10}\b', text)
            if long_nums:
                data["invoice_no"] = long_nums[0]

        # Date pattern
        date_match = re.search(r'(?:Memo\s*Date|Invoice\s*Date)\s*[:\s]*([0-9]{1,2}-[A-Za-z]{3}-[0-9]{2,4})', text, re.IGNORECASE)
        if date_match:
            data["invoice_date"] = date_match.group(1).strip()

        # Period
        period_match = re.search(r'Period\s*[:\s]*([A-Za-z]{3}-[0-9]{2,4})', text, re.IGNORECASE)
        if period_match:
            data["billing_period"] = period_match.group(1).strip()

        # Number of sites
        sites_match = re.search(r'(?:No\s*of\s*Sites|Sites)\s*[:\s]*([0-9]{1,4})', text, re.IGNORECASE)
        if sites_match:
            try:
                data["no_of_sites"] = int(sites_match.group(1))
            except ValueError:
                data["no_of_sites"] = 0

        # Taxable amount
        taxable_match = re.search(r'Total\s*Taxable\s*Amount\s*[:\s]*([0-9,]+\.[0-9]{2})', text, re.IGNORECASE)
        if taxable_match:
            data["taxable_amount"] = float(taxable_match.group(1).replace(",", ""))

        # GST amount
        cgst_match = re.search(r'Central\s*GST[^\n]*?([0-9,]+\.[0-9]{2})', text, re.IGNORECASE)
        if cgst_match:
            data["cgst"] = float(cgst_match.group(1).replace(",", ""))

        sgst_match = re.search(r'State\s*GST[^\n]*?([0-9,]+\.[0-9]{2})', text, re.IGNORECASE)
        if sgst_match:
            data["sgst"] = float(sgst_match.group(1).replace(",", ""))

        # Total amount
        total_match = re.search(r'(?:Total\s*Amount|Amount\s*in\s*words)[^\n]*?([0-9,]+\.[0-9]{2})', text, re.IGNORECASE)
        if total_match:
            data["total_amount"] = float(total_match.group(1).replace(",", ""))

        return data

    @staticmethod
    def parse_site_annexure_excel(file_path_or_buffer) -> pd.DataFrame:
        """
        Load Site Annexure sheet from Excel file (.xlsx or .xls).
        Normalizes column names to standard BSNL audit headers.
        """
        xl = pd.ExcelFile(file_path_or_buffer)
        sheet_names = xl.sheet_names

        # Pick target sheet
        target_sheet = sheet_names[0]
        for name in sheet_names:
            if "FEB-26" in name.upper() or "ANNEXURE" in name.upper() or "SITE" in name.upper():
                target_sheet = name
                break

        # Read the raw sheet to find the header row
        df_raw = xl.parse(target_sheet, header=None)
        
        # Look for header row that contains 'Global ID' or 'Site Name' or 'SSA'
        header_row_idx = 1
        for idx in range(min(10, len(df_raw))):
            row_vals = [str(v).lower() for v in df_raw.iloc[idx].values if pd.notna(v)]
            if any("global" in v for v in row_vals) or any("site id" in v for v in row_vals) or any("ssa" in v for v in row_vals):
                header_row_idx = idx
                break

        df = xl.parse(target_sheet, skiprows=header_row_idx)
        df = InvoiceParser._normalize_columns(df)
        return df

    @classmethod
    def parse_invoice_stream(cls, file_or_text) -> Dict[str, pd.DataFrame]:
        """
        Parses the new standard invoice format (CSV, Excel, or raw text) containing:
        1. Main billing line items (GLOBAL_SITE_ID, Applied MSA Rate, Total Billing, Invoice no)
        2. Deployment Penalty section (Days_SO to RFI, Deployment penalty)
        """
        import io
        if hasattr(file_or_text, "read"):
            raw_content = file_or_text.read()
            if isinstance(raw_content, bytes):
                text = raw_content.decode("utf-8", errors="ignore")
            else:
                text = str(raw_content)
        elif os.path.exists(str(file_or_text)):
            with open(str(file_or_text), "r", encoding="utf-8", errors="ignore") as f:
                text = f.read()
        else:
            text = str(file_or_text)

        lines = text.split('\n')
        billing_lines = []
        penalty_lines = []
        
        current_target = None
        for line in lines:
            line_clean = line.strip()
            if not line_clean:
                continue
            if 'GLOBAL_SITE_ID' in line_clean and 'Applied MSA Rate' in line_clean:
                current_target = 'billing'
                billing_lines.append(line_clean)
                continue
            elif 'GLOBAL_SITE_ID' in line_clean and 'Deployment penalty' in line_clean:
                current_target = 'penalty'
                penalty_lines.append(line_clean)
                continue
            elif 'IP hired sites Down' in line_clean:
                break
                
            if current_target == 'billing':
                if line_clean.startswith(',') and not any(c.isalnum() for c in line_clean):
                    continue
                billing_lines.append(line_clean)
            elif current_target == 'penalty':
                if line_clean.startswith(',') and not any(c.isalnum() for c in line_clean[:10]):
                    continue
                penalty_lines.append(line_clean)
                
        df_billing = pd.read_csv(io.StringIO('\n'.join(billing_lines))) if billing_lines else pd.DataFrame()
        df_penalty = pd.read_csv(io.StringIO('\n'.join(penalty_lines))) if penalty_lines else pd.DataFrame()
        
        # Normalize billing columns
        if not df_billing.empty:
            df_billing = cls._normalize_columns(df_billing)
            
        return {
            'billing_df': df_billing,
            'penalty_df': df_penalty
        }

    @staticmethod
    def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
        """Clean and map messy column headers to standardized unique names."""
        clean_cols = {}
        used_names = set()

        for col in df.columns:
            c_str = str(col).strip()
            c_lower = c_str.lower()
            name = None

            if "global" in c_lower and "id" in c_lower:
                name = "global_id"
            elif "vendor" in c_lower and "id" in c_lower:
                name = "vendor_site_id"
            elif "atc site" in c_lower or "altius site" in c_lower:
                name = "vendor_site_id"
            elif "bsnl" in c_lower and "id" in c_lower:
                name = "bsnl_site_id"
            elif "site" in c_lower and "name" in c_lower:
                if "site_name" not in used_names:
                    name = "site_name"
                else:
                    name = "bsnl_site_name"
            elif "ssa" in c_lower and "type" not in c_lower:
                name = "ssa"
            elif "tenancy" in c_lower:
                if "claimed" in c_lower:
                    name = "claimed_tenancy"
                elif "working" in c_lower or "tenancy" not in used_names:
                    name = "working_tenancy"
                else:
                    name = "claimed_tenancy"
            elif "tower" in c_lower and "type" in c_lower:
                name = "tower_type"
            elif "bts type" in c_lower:
                if "claimed" in c_lower:
                    name = "claimed_bts_type"
                elif "rental" in c_lower or "bts_type" not in used_names:
                    name = "bts_type"
            elif "city" in c_lower or "class" in c_lower:
                if "claimed" in c_lower:
                    name = "claimed_city_class"
                elif "city_class" not in used_names:
                    name = "city_class"
            elif "outage" in c_lower:
                if "min" in c_lower:
                    name = "outage_minutes"
                elif "rs" in c_lower or "calculated" in c_lower:
                    name = "calculated_outage_rs"
            elif "dg run" in c_lower or "dg running" in c_lower:
                if "claimed" in c_lower:
                    name = "claimed_dg_run_hours"
                elif "dg_run_hours" not in used_names:
                    name = "dg_run_hours"
            elif "dg cph" in c_lower:
                name = "dg_cph"
            elif "diesel" in c_lower and ("rate" in c_lower or "cost" in c_lower or "ltr" in c_lower):
                name = "diesel_rate"
            elif "eb rate" in c_lower:
                name = "eb_rate"
            elif "eb cph" in c_lower:
                name = "eb_cph"
            elif "invoice no" in c_lower:
                name = "invoice_no"
            elif "total billing" in c_lower or "claimed energy" in c_lower or "claimed ip fee" in c_lower or "total amount before tax" in c_lower:
                name = "claimed_amount"
            elif "total amount after tax" in c_lower:
                name = "claimed_with_tax"
            elif "applied msa rate" in c_lower or "net applied rate" in c_lower:
                name = "msa_rate"
            elif "opco site id" in c_lower:
                name = "bsnl_site_id"
            elif "site type" in c_lower:
                name = "tower_type"
            elif "remark" in c_lower:
                name = "field_remark"

            if name and name not in used_names:
                clean_cols[col] = name
                used_names.add(name)

        df = df.rename(columns=clean_cols)

        # Remove duplicate column names if any exist
        df = df.loc[:, ~df.columns.duplicated()].copy()

        # Clean global_id
        if "global_id" in df.columns:
            df["global_id"] = pd.to_numeric(df["global_id"], errors="coerce")
            df = df.dropna(subset=["global_id"])
            df["global_id"] = df["global_id"].astype(int).astype(str)

        return df
