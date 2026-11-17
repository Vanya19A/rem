import sys, os
sys.path.insert(0, os.path.abspath('.'))
from engine.invoice_parser import InvoiceParser
from engine.outage_parser import OutageParser

sample_inv = """Srno,GLOBAL_SITE_ID,Site ID,SAP ID,OPCO Site ID,Site Name,City Class,Site Type,No.of Tenants,OPCO Name,BTS Type,Date of Issue of SO,Date of RFI,4G Integration Date,Biiling Start Date,Period Billing From,Period Billing To,Applied MSA Rate,PRORATED_STANDARD_IPFEE,Escalation Amount,Net Applied Rate,Antenna Markup-EA,Power Markup,BTS Markup,Total Loading Mark up,Total Loading Mark up with Esc,Total Billing,PREVIOUS_BILLED_IP_AMOUNT,Total Amount before Tax,Tx,Total Amount after Tax,Invoice no
1,503751,GU/SU/BS/0097,GU/SU/BS/0097,SUR395,"PUNGRAM SURAT",B1,RTT,4,"VODAFONE 10/17/2015",OD,2011-03-09,2011-09-28,2024-10-21,2011-09-28,2026-04-01,2026-04-30,13343,13343,0,13343,650,3000,3750,7400,7400,20743,0,20743,3733.74,24476.74,607249101046925
Srno,GLOBAL_SITE_ID,Site ID,SAP ID,OPCO Site ID,Site Name,Site Type,Date of Issue of SO,Date of RFI,Days_SO to RFI,Deployment penalty
1,530882,GU/PA/TT/12982,GUJPAT2170,"MEH372","SHANKHESHVAR",GBT,2018-07-03,2018-12-24,174,40000"""

inv_res = InvoiceParser.parse_invoice_stream(sample_inv)
print("=== INVOICE PARSER RESULTS ===")
print("Billing sites:", len(inv_res['billing_df']))
print("Columns normalized:", inv_res['billing_df'].columns.tolist()[:8])
print("Deployment penalty records:", len(inv_res['penalty_df']))
print(inv_res['penalty_df'][['GLOBAL_SITE_ID', 'OPCO Site ID', 'Days_SO to RFI', 'Deployment penalty']])

sample_out = """IP hired sites Down Time and AC status Report,,,,,,,,,,
Name of SSA : Surat,,,,,Month : Apr-2026,,,,,
Name Of IP : Signotox,,,,,,,,,,
Sr No,IP ID,OMCR ID,Site Name,ID/OD,Down From,Down To,Duration (in Hr),Fault Reason,AC status,AC Remark
1.0,GURPUSRT0015,SRT610,SRT610_MahatmaIndstate,OD,Nil,Nil,Nil,Nil,NA,NA
IP hired sites Down Time and AC status Report,,,,,,,,,,
Name of SSA : HIMMATNAGAR,,,,,Month :Apr-2026,,,,,
Name Of IP : Summit Digitel Infrastructure Pvt. Ltd,,,,,,,,,,
Sr No,IP ID,OMCR ID,Site Name,ID/OD,Down From,Down To,Duration (in Hr),Fault Reason,AC status,AC Remark
1.0,I-GJ-HGRA-ENB-0005,HMT256,HMT256_Panpur Patia,OD,NA,NA,15,NIL,NA,BTS Down Time Report
IP hired sites Down Time and AC status Report,,,,,,,,,,
Name of SSA : Junagadh,,,,,Month : April-2026,,,,,
Name Of IP : Tower Vision,,,,,,,,,,
Sr No,IP ID,OMCR ID,Site Name,ID/OD,Down From,Down To,Duration (in Hr),Fault Reason,AC status,AC Remark
1.0, VDGJPBD0017,JND314,JND314_Gitanagar,OD,28/04/2026 13:37:00 PM,28/04/2026 18:15:00 PM,4Hrs 38 Minutes,Owner Issue,NA,N/A"""

out_df = OutageParser.parse_multi_block_outage(sample_out)
print("\n=== OUTAGE PARSER RESULTS ===")
print("Total outage rows parsed:", len(out_df))
print(out_df[['operator', 'ssa', 'vendor_site_id', 'bsnl_site_id', 'duration_hours', 'outage_minutes', 'fault_reason']])
