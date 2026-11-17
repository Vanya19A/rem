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

    @classmethod
    def parse_multi_block_outage(cls, file_or_text) -> pd.DataFrame:
        """
        Parses the multi-SSA, multi-operator outage sign-off format (CSV/Excel/Text).
        Extracts every downtime entry across all operators (Signotox, Summit, Indus, GTL, Tower Vision, Ascend, ATC) and all SSAs.
        """
        import io
        import os

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

        lines = [l.strip() for l in text.split('\n') if l.strip()]
        records = []
        
        current_ssa = ''
        current_month = ''
        current_ip = ''
        in_table = False
        headers = []
        
        for line in lines:
            if 'IP hired sites Down' in line or 'Report For the Month' in line or 'O/o DE (CMTS)' in line:
                in_table = False
                continue
            
            m_ssa = re.search(r'Name of SSA\s*:\s*([^,]+)', line, re.I)
            m_month = re.search(r'Month\s*:\s*([^,]+)', line, re.I)
            if m_ssa:
                current_ssa = m_ssa.group(1).strip()
            if m_month:
                current_month = m_month.group(1).strip()
                
            m_ip = re.search(r'Name Of IP\s*:\s*([^,]+)', line, re.I)
            if m_ip:
                current_ip = m_ip.group(1).strip()
                continue
                
            if ('IP ID' in line or 'OMCR ID' in line or 'Site ID' in line) and ('Site Name' in line or 'Duration' in line or 'Down' in line):
                in_table = True
                headers = [c.strip().strip('"').strip() for c in line.split(',')]
                continue
                
            if in_table:
                parts = [p.strip().strip('"').strip() for p in line.split(',')]
                if len(parts) >= 4 and any(parts[1:5]):
                    row_dict = {
                        'operator': current_ip,
                        'ssa': current_ssa,
                        'month': current_month
                    }
                    for i, h in enumerate(headers):
                        if i < len(parts):
                            h_clean = re.sub(r'\s+', ' ', h).strip()
                            row_dict[h_clean] = parts[i]
                    
                    dur_raw = row_dict.get('Duration (in Hr)') or row_dict.get('Duration (in minutes)') or row_dict.get('Down Time in Hrs') or '0'
                    dur_hrs = cls._parse_duration_val(dur_raw)
                    if 'minutes' in str(dur_raw).lower():
                        dur_mins = dur_hrs
                        dur_hrs = dur_mins / 60.0
                    else:
                        dur_mins = dur_hrs * 60.0
                    
                    row_dict['duration_hours'] = round(dur_hrs, 2)
                    row_dict['outage_minutes'] = round(dur_mins, 2)
                    
                    ip_id = row_dict.get('IP ID') or row_dict.get('Site ID') or row_dict.get('SL.No.') or ''
                    omcr_id = row_dict.get('OMCR ID') or row_dict.get('Site ID') or ''
                    row_dict['vendor_site_id'] = str(ip_id).strip()
                    row_dict['bsnl_site_id'] = str(omcr_id).strip()
                    row_dict['fault_reason'] = row_dict.get('Fault Reason') or row_dict.get('Reason of Fault') or ''
                    row_dict['ac_status'] = row_dict.get('AC status (For ID sites OK/NOK) For OD Sites NA') or row_dict.get('AC status') or ''
                    row_dict['officer_remarks'] = row_dict.get('Remarks') or row_dict.get('AC Remark / Duration of Fault Days') or ''
                    
                    records.append(row_dict)
                    
        return pd.DataFrame(records)

    @classmethod
    def _parse_duration_val(cls, val_str) -> float:
        if not val_str or str(val_str).strip().upper() in ['NIL', 'NA', 'N/A', '-', '']:
            return 0.0
        s = str(val_str).strip()
        m_hm = re.search(r'(\d+)\s*(?:hrs|hr|hours)?\s*(\d+)?\s*(?:min|minutes|mins)?', s, re.I)
        try:
            return float(s)
        except:
            pass
        if m_hm:
            h = float(m_hm.group(1)) if m_hm.group(1) else 0.0
            m = float(m_hm.group(2)) if m_hm.group(2) else 0.0
            return h + m / 60.0
        return 0.0
