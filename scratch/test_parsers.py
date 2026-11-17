import re
import io
from typing import Dict, List, Any, Optional
import pandas as pd

def parse_duration_to_hours(val_str):
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

def parse_multi_block_outage_text(text: str) -> pd.DataFrame:
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
        
        # Check SSA & Month
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
            
        # Detect table headers
        if ('IP ID' in line or 'OMCR ID' in line or 'Site ID' in line) and ('Site Name' in line or 'Duration' in line or 'Down' in line):
            in_table = True
            headers = [c.strip().strip('"').strip() for c in line.split(',')]
            continue
            
        if in_table:
            parts = [p.strip().strip('"').strip() for p in line.split(',')]
            if len(parts) >= 4 and any(parts[1:5]):
                # Normalize column keys
                row_dict = {
                    'operator': current_ip,
                    'ssa': current_ssa,
                    'month': current_month
                }
                for i, h in enumerate(headers):
                    if i < len(parts):
                        h_clean = re.sub(r'\s+', ' ', h).strip()
                        row_dict[h_clean] = parts[i]
                
                # Extract duration
                dur_raw = row_dict.get('Duration (in Hr)') or row_dict.get('Duration (in minutes)') or row_dict.get('Down Time in Hrs') or '0'
                dur_hrs = parse_duration_to_hours(dur_raw)
                if 'minutes' in str(dur_raw).lower():
                    dur_mins = dur_hrs
                    dur_hrs = dur_mins / 60.0
                else:
                    dur_mins = dur_hrs * 60.0
                
                row_dict['duration_hours'] = round(dur_hrs, 2)
                row_dict['outage_minutes'] = round(dur_mins, 2)
                records.append(row_dict)
                
    return pd.DataFrame(records)

sample_data = '''IP hired sites Down Time and AC status Report,,,,,,,,,,
Name of SSA : Surat,,,,,Month : Apr-2026,,,,,
Name Of IP : Signotox,,,,,,,,,,
Sr No,IP ID,OMCR ID,Site Name,ID/OD,Down From,Down To,Duration (in Hr),Fault Reason,AC status,AC Remark
1.0,GURPUSRT0015,SRT610,SRT610_MahatmaIndstate,OD,Nil,Nil,Nil,Nil,NA,NA
2.0,GURPUSRT0062,SRT621,SRT621_Kevalkunj,OD,Nil,Nil,Nil,Nil,NA,NA
IP hired sites Down Time and AC status Report,,,,,,,,,,
Name of SSA : HIMMATNAGAR,,,,,Month :Apr-2026,,,,,
Name Of IP : Summit Digitel Infrastructure Pvt. Ltd,,,,,,,,,,
Sr No,IP ID,OMCR ID,Site Name,ID/OD,Down From,Down To,Duration (in Hr),Fault Reason,AC status,AC Remark
1.0,I-GJ-HGRA-ENB-0005,HMT256,HMT256_Panpur Patia,OD,NA,NA,15,NIL,NA,BTS Down Time Report'''

df = parse_multi_block_outage_text(sample_data)
print('Parsed records count:', len(df))
print(df[['operator', 'ssa', 'IP ID', 'OMCR ID', 'duration_hours', 'outage_minutes']].to_string(index=False))

def parse_incoming_invoice_text(text: str) -> Dict[str, pd.DataFrame]:
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
            # Check if line has columns
            if line_clean.startswith(',') and not any(c.isalnum() for c in line_clean):
                continue
            billing_lines.append(line_clean)
        elif current_target == 'penalty':
            if line_clean.startswith(',') and not any(c.isalnum() for c in line_clean[:10]):
                continue
            penalty_lines.append(line_clean)
            
    df_billing = pd.read_csv(io.StringIO('\n'.join(billing_lines))) if billing_lines else pd.DataFrame()
    df_penalty = pd.read_csv(io.StringIO('\n'.join(penalty_lines))) if penalty_lines else pd.DataFrame()
    
    return {'billing_df': df_billing, 'penalty_df': df_penalty}

sample_inv = '''Srno,GLOBAL_SITE_ID,Site ID,SAP ID,OPCO Site ID,Site Name,City Class,Site Type,No.of Tenants,OPCO Name,BTS Type,Date of Issue of SO,Date of RFI,4G Integration Date,Biiling Start Date,Period Billing From,Period Billing To,Applied MSA Rate,PRORATED_STANDARD_IPFEE,Escalation Amount,Net Applied Rate,Antenna Markup-EA,Power Markup,BTS Markup,Total Loading Mark up,Total Loading Mark up with Esc,Total Billing,PREVIOUS_BILLED_IP_AMOUNT,Total Amount before Tax,Tx,Total Amount after Tax,Invoice no
1,503751,GU/SU/BS/0097,GU/SU/BS/0097,SUR395,"PUNGRAM SURAT",B1,RTT,4,"VODAFONE 10/17/2015",OD,2011-03-09,2011-09-28,2024-10-21,2011-09-28,2026-04-01,2026-04-30,13343,13343,0,13343,650,3000,3750,7400,7400,20743,0,20743,3733.74,24476.74,607249101046925
Srno,GLOBAL_SITE_ID,Site ID,SAP ID,OPCO Site ID,Site Name,Site Type,Date of Issue of SO,Date of RFI,Days_SO to RFI,Deployment penalty
1,530882,GU/PA/TT/12982,GUJPAT2170,"MEH372","SHANKHESHVAR",GBT,2018-07-03,2018-12-24,174,40000'''

res = parse_incoming_invoice_text(sample_inv)
print('Billing sites parsed:', len(res['billing_df']))
print('Penalty sites parsed:', len(res['penalty_df']))
print(res['billing_df'][['GLOBAL_SITE_ID', 'OPCO Site ID', 'Site Name', 'Total Billing', 'Invoice no']])
print(res['penalty_df'])
