"""
Outage Sign-Off Parser for BSNL Field Reports
Extracts SSA-wise outage sign-offs, outage minutes, tenant counts, and officer remarks.
"""

import re
from typing import Dict, List, Any, Optional
import pandas as pd
try:
    import fitz  # Standard PyMuPDF
except Exception:
    try:
        import pymupdf as fitz
    except Exception:
        fitz = None


class OutageParser:
    def __init__(self):
        self._ocr = None

    def _get_ocr(self):
        if self._ocr is None:
            from rapidocr_onnxruntime import RapidOCR
            self._ocr = RapidOCR()
        return self._ocr

    def parse_outage_pdf(self, pdf_path: str, progress_callback=None) -> pd.DataFrame:
        """
        Parse Outage Sign-Off sheets across all pages of a PDF.
        """
        if fitz is None:
            raise ImportError("PyMuPDF is not properly installed. Please run: pip install PyMuPDF==1.24.14")
        doc = fitz.open(pdf_path)
        records = []
        total_pages = len(doc)

        for page_idx in range(total_pages):
            if progress_callback:
                progress_callback(page_idx + 1, total_pages)

            page = doc[page_idx]
            text = page.get_text()

            # If page contains little/no text, OCR the page
            lines = []
            if len(text.strip()) < 50:
                pix = page.get_pixmap(dpi=150)
                ocr_engine = self._get_ocr()
                ocr_res, _ = ocr_engine(pix.tobytes("png"))
                if ocr_res:
                    lines = [r[1] for r in ocr_res]
            else:
                lines = [l.strip() for l in text.split("\n") if l.strip()]

            # Check if this page is an outage sign-off
            full_text = " ".join(lines).upper()
            if "OUTAGE SIGN OFF" in full_text or "OUTAGE IN MINUTES" in full_text or "TOTAL DOWN TIME" in full_text:
                page_records = self._extract_outage_rows(lines, page_idx + 1)
                records.extend(page_records)

        if not records:
            return pd.DataFrame()

        df = pd.DataFrame(records)
        # Drop duplicates by global_id, keeping the one with higher outage
        if "global_id" in df.columns:
            df = df.sort_values(by="outage_minutes", ascending=False).drop_duplicates(subset=["global_id"])
        return df

    def _extract_outage_rows(self, lines: List[str], page_num: int) -> List[Dict[str, Any]]:
        """Extract individual site rows and remarks from page lines."""
        rows = []
        detected_ssa = ""
        remarks = []

        # Find SSA
        for line in lines:
            ssa_m = re.search(r'SSA\s*[:\-\s]*([A-Za-z]+)', line, re.IGNORECASE)
            if ssa_m:
                detected_ssa = ssa_m.group(1).upper()
            elif "GUNA" in line.upper():
                detected_ssa = "GUNA"
            elif "GWALIOR" in line.upper():
                detected_ssa = "GWALIOR"
            elif "BHOPAL" in line.upper():
                detected_ssa = "BHOPAL"
            elif "INDORE" in line.upper():
                detected_ssa = "INDORE"

            # Check for remarks written by SDE / JTO
            if any(k in line.lower() for k in ["theft", "down", "deleted", "working", "power off", "fault", "cable"]):
                remarks.append(line)

        combined_remark = "; ".join(remarks) if remarks else ""

        # Scan for Global IDs (5-6 digit numbers like 557737, 460848, 500882)
        for i, line in enumerate(lines):
            # Find 6-digit site ID
            id_match = re.search(r'\b(4\d{5}|5\d{5}|6\d{5})\b', line)
            if id_match:
                gid = id_match.group(1)

                # Look forward in next 1-4 lines for Outage in Minutes and Tenancy
                outage_min = 0.0
                tenancy = None
                site_name = ""

                # Look around this line for numbers that represent outage or tenancy
                context_window = lines[max(0, i - 1): min(len(lines), i + 6)]
                context_text = " ".join(context_window)

                # Extract outage minutes (number directly associated or 'NIL' / '0')
                min_match = re.search(r'\b(\d{1,5})\s*(?:mins?|minutes?)\b', context_text, re.IGNORECASE)
                if min_match:
                    outage_min = float(min_match.group(1))
                else:
                    # check for standalone number at the end of line
                    nums = re.findall(r'\b(\d{1,5})\b', line)
                    if len(nums) >= 2:
                        # usually last number is outage minutes
                        candidate = float(nums[-1])
                        if candidate != float(gid) and candidate <= 44640:
                            outage_min = candidate

                # Extract tenancy (1, 2, 3, or 4)
                ten_matches = re.findall(r'\b([1-4])\b', context_text)
                if ten_matches:
                    tenancy = int(ten_matches[-1])

                rows.append({
                    "page": page_num,
                    "ssa": detected_ssa,
                    "global_id": str(gid),
                    "outage_minutes": float(outage_min),
                    "verified_tenancy": tenancy,
                    "field_remark": combined_remark
                })

        return rows
