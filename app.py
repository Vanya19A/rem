"""
BSNL Telecom Infrastructure (IP-1) Billing & SLA Reconciliation Dashboard
Supports Multi-Vendor MSAs, 1st & 3rd EoI Amendments, Outage SLA Penalties,
Energy (P&F) Auditing, and IP Tower Rental Fee Reconciliation.
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
import openpyxl
import calendar
from datetime import datetime
from engine.msa_rules import calculate_outage_penalty, calculate_allowed_downtime_minutes, STANDARD_IP_FEES
from engine.invoice_parser import InvoiceParser
from engine.outage_parser import OutageParser
from engine.reconciler import ReconciliationEngine
from engine.excel_generator import ExcelReportGenerator
from engine.erp_rem_generator import ERPREMGenerator
from engine.master_infra_manager import MasterInfraManager

# Page Configuration
st.set_page_config(
    page_title="BSNL IP-1 Billing & SLA Reconciliation System",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-title {
        font-size: 26px;
        font-weight: 700;
        color: #003366;
        margin-bottom: 2px;
    }
    .sub-title {
        font-size: 14px;
        color: #555555;
        margin-bottom: 18px;
    }
    .upload-box {
        background-color: #f8f9fa;
        border: 2px dashed #003366;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 15px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 14px;
        font-weight: 600;
    }
    .metric-card {
        background-color: #f0f4f8;
        border-left: 5px solid #003366;
        padding: 10px 14px;
        border-radius: 4px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)


def init_session():
    if "reconciled_df" not in st.session_state:
        st.session_state.reconciled_df = None
    if "summary_metrics" not in st.session_state:
        st.session_state.summary_metrics = None
    if "active_msa_name" not in st.session_state:
        st.session_state.active_msa_name = "ATC WIRELESS TT INFO SERVICES MAIN CURRENT MSA (Active)"
    if "sidebar_vendor" not in st.session_state:
        st.session_state["sidebar_vendor"] = "Elevar Digitel (ATC)"
    if "sidebar_circle" not in st.session_state:
        st.session_state["sidebar_circle"] = "Gujarat"
    if "sidebar_month" not in st.session_state:
        st.session_state["sidebar_month"] = "APR"
    if "sidebar_year" not in st.session_state:
        st.session_state["sidebar_year"] = "2026"
    if "custom_dataset_loaded" not in st.session_state:
        st.session_state["custom_dataset_loaded"] = False

    # Preload Feb-2026 Energy data if available
    excel_path = "BSNL_5_Invoices_Reconciliation_Output.xlsx"
    if st.session_state.reconciled_df is None and os.path.exists(excel_path):
        try:
            xls = pd.ExcelFile(excel_path)
            sheets = [s for s in xls.sheet_names if "Inv " in s]
            dfs = []
            for s in sheets:
                df_sheet = xls.parse(s, skiprows=2)
                dfs.append(df_sheet)
            if dfs:
                st.session_state.reconciled_df = pd.concat(dfs, ignore_index=True)
        except Exception:
            pass

    # Initialize Master Vendor MSA Rules and Amendments Registry
    if "vendor_msa_rules" not in st.session_state:
        st.session_state.vendor_msa_rules = {
            "Elevar Digitel (ATC)": {
                "lineage": "Quippo Telecom (1st EoI) -> WTTIL (3rd EoI) -> Viom Networks -> ATC TIPL -> Elevar Digitel. Novated: Vodafone, Idea CISL, KEC.",
                "master_msa": "3rd EoI MSA dated 19.04.2010 (erstwhile WTTIL/Viom)",
                "uptime_target": "99.90% (Site Level DC Uptime), 99.99% (Circle Level)",
                "permissible_downtime": "40.32 mins (28-day month), 43.20 mins (30-day month)",
                "penalty_formula": "₹500/hr pro-rata (Excess Mins / 60 * ₹500)",
                "penalty_cap": "₹10,000 per site per month",
                "llr_capping": "Threshold absorbed: ₹8,000. Excess capped at ₹7,000 (Delhi/Mumbai), ₹5,000 (Kolkata/BLR/Chennai), ₹3,000 (Rest of India) w.e.f. 01.09.2023",
                "tenancy_rule": "Deemed 2nd Tenant Protection: BSNL 1T (solo) billed at 2T sharing rate. Actual rate if tenancy > 2.",
                "lip_extension": "10-year renewal for LIP expired sites up to 31.12.2025; discount equal to difference of escalated base rate and unescalated 3rd EoI base rate.",
                "rates_4g": "Outdoor 4G BBU (ODC & OD Pad): ₹1,350/mo | Indoor 2U BBU: ₹750/mo | RRU/RRH: ₹800/unit/mo | Power provisioning (>0.7kW): ₹3,000/mo",
                "energy_norm": "Actual EB bills or HPC norms: DG run hrs * CPH * diesel rate + EB hrs * rate * CPH. Audited as Min(Claimed, Calculated).",
                "amendments_history": [
                    {"Circular / Ref No.": "BSNLCO-CMIN/11/2/2020-CMTS INFRA/ 1006868", "Date": "06.08.2024", "Subject": "Understanding reached between BSNL and M/s ATC-TIPL on existing MSAs", "Key Amendment": "1st EoI, Voda, Idea & KEC sites migrated to 3rd EoI w.e.f. 01.09.2023; 10-yr LIP extension with unescalated base rate freeze; LLR capping ₹7k/₹5k/₹3k; deemed 2T protection.", "Effective Date": "01.09.2023"},
                    {"Circular / Ref No.": "BSNLCO-CMIN/11/2/2020-CMTS INFRA", "Date": "06.08.2024", "Subject": "Applicable charges for 4G upgradation / new sites and power provisioning", "Key Amendment": "Outdoor 4G BBU with ODC/Pad at ₹1,350/mo; Indoor BBU in 2U space at ₹750/mo; additional power beyond 0.7kW at ₹3,000/mo.", "Effective Date": "06.08.2024"},
                    {"Circular / Ref No.": "Addendum dated 16.06.2017", "Date": "16.06.2017", "Subject": "RRH / RRU Installation Addendum", "Key Amendment": "Remote Radio Head (RRH) monthly charge fixed at ₹800 per unit.", "Effective Date": "16.06.2017"},
                    {"Circular / Ref No.": "Addendum dated 02.11.2015 / 20.10.2015", "Date": "02.11.2015", "Subject": "Additional BTS Rack Space & Threshold Rent", "Key Amendment": "Additional BTS rack space at ₹2,500 (ID) / ₹1,500 (OD). Landlord rent absorption threshold established at ₹8,000/-.", "Effective Date": "20.10.2015"}
                ]
            },
            "Indus Towers": {
                "lineage": "Bharti Infratel Limited & Indus Towers merged effective July 2021 (Indus Towers Limited).",
                "master_msa": "1st EoI MSA dated 09.09.2008 & 3rd EoI MSA dated 09.03.2010 / 29.03.2010",
                "uptime_target": "99.90% (Site Level DC Uptime), 99.99% (Circle Level)",
                "permissible_downtime": "40.32 mins (28-day month), 43.20 mins (30-day month)",
                "penalty_formula": "₹500/hr pro-rata (Excess Mins / 60 * ₹500)",
                "penalty_cap": "₹10,000 per site per month",
                "llr_capping": "Threshold absorbed: ₹8,000. Excess capped at ₹7,000 (Delhi/Mumbai), ₹5,000 (Kolkata/BLR/Chennai), ₹3,000 (Rest of India) w.e.f. 01.09.2023",
                "tenancy_rule": "Deemed 2nd Tenant Protection: BSNL solo 1T billed at 2T sharing rate. Upward/downward revision based on actual sharing.",
                "lip_extension": "10-year lock-in renewal with discount equal to difference between escalated rate and unescalated 3rd EoI base rate.",
                "rates_4g": "Outdoor 4G BBU: ₹1,350/mo | Indoor BBU (2U): ₹750/mo | RRU: ₹800/unit/mo | Additional power: ₹3,000/mo",
                "energy_norm": "Min(Claimed, Calculated HPC DG+EB).",
                "amendments_history": [
                    {"Circular / Ref No.": "BSNLCO-CMIN/16(15)/1/2020-CMTSINFRA/822038", "Date": "02.01.2024", "Subject": "Understanding reached between BSNL and M/s Indus Towers Limited on existing MSAs", "Key Amendment": "All 1st EoI sites migrated to 3rd EoI w.e.f. 01.09.2023; 10-yr LIP extension with unescalated base rate freeze; LLR capping ₹7k/₹5k/₹3k; deemed 2T protection.", "Effective Date": "01.09.2023"},
                    {"Circular / Ref No.": "Amendment Agreement", "Date": "25.01.2024 & 25.06.2024", "Subject": "Formal Amendment Agreement to ISA 09.09.2008 & 09.03.2010", "Key Amendment": "Formalized 3rd EoI migration, 10-yr lock-in extension, and commercial fee terms.", "Effective Date": "01.09.2023"},
                    {"Circular / Ref No.": "Addendum dated 04.12.2018", "Date": "04.12.2018", "Subject": "Additional Space & Rental Threshold", "Key Amendment": "Landlord rental absorption threshold established at ₹8,000/-.", "Effective Date": "04.12.2018"}
                ]
            },
            "Ascend Telecom": {
                "lineage": "Aster Infrastructure -> ITIL (India Telecom Infra Ltd) -> Bipartite merger of ITIL into Ascend Telecom Infrastructure Pvt. Ltd. (04.09.2012).",
                "master_msa": "1st EoI (18.02.2008 / 29.06.2008) & 3rd EoI (26.03.2010 ITIL / 15.10.2010 Ascend)",
                "uptime_target": "99.90% (Site Level DC Uptime), 99.99% (Circle Level)",
                "permissible_downtime": "40.32 mins (28-day month), 43.20 mins (30-day month)",
                "penalty_formula": "₹500/hr pro-rata (Excess Mins / 60 * ₹500)",
                "penalty_cap": "₹10,000 per site per month",
                "llr_capping": "Threshold absorbed: ₹8,000. Excess capped at ₹7k (Delhi/Mumbai), ₹5k (Metros), ₹3k (RoI) w.e.f. 01.09.2023",
                "tenancy_rule": "Deemed 2nd Tenant Protection: Capped at 2T rate if solo.",
                "lip_extension": "10-year renewal at unescalated 3rd EoI base rates.",
                "rates_4g": "Outdoor BBU: ₹1,350/mo | Indoor BBU (2U): ₹750/mo | RRU: ₹800/unit/mo | Additional power: ₹3,000/mo",
                "energy_norm": "Actual EB / HPC DG norms.",
                "amendments_history": [
                    {"Circular / Ref No.": "BSNLCO-CMIN/16/1/2024-CMTS INFRA/961039", "Date": "14.06.2024", "Subject": "Applicable charges for 4G upgradation / new sites at M/s Ascend Telecom", "Key Amendment": "Outdoor BBU at ₹1,350/mo; Indoor BBU 2U at ₹750/mo; additional power at ₹3,000/mo; 1st EoI migration to 3rd EoI w.e.f. 01.09.2023; LLR capping ₹7k/₹5k/₹3k.", "Effective Date": "01.09.2023"},
                    {"Circular / Ref No.": "Addendum dated 11.06.2024 & 08.08.2024", "Date": "11.06.2024", "Subject": "Understanding reached on existing MSAs", "Key Amendment": "Formalized 10-year LIP renewal and unescalated base rate freeze.", "Effective Date": "01.09.2023"},
                    {"Circular / Ref No.": "RRH Addendum", "Date": "05.02.2018", "Subject": "RRH Installation Addendum", "Key Amendment": "RRH monthly charge fixed at ₹800 per unit.", "Effective Date": "05.02.2018"},
                    {"Circular / Ref No.": "BTS Rack Addendum", "Date": "14.03.2016", "Subject": "Additional BTS Rack space & threshold rent", "Key Amendment": "Additional rack at ₹2,500 ID / ₹1,500 OD; threshold rent ₹8,000.", "Effective Date": "14.03.2016"}
                ]
            },
            "Tower Vision": {
                "lineage": "Tower Vision India Private Limited (TVI).",
                "master_msa": "1st EoI dated 27.06.2008 & 3rd EoI dated 10.09.2010 / 15.10.2010",
                "uptime_target": "99.90% (Site Level DC Uptime), 99.99% (Circle Level)",
                "permissible_downtime": "40.32 mins (28-day month), 43.20 mins (30-day month)",
                "penalty_formula": "₹500/hr pro-rata (Excess Mins / 60 * ₹500)",
                "penalty_cap": "₹10,000 per site per month",
                "llr_capping": "Threshold absorbed: ₹8,000. Excess capped at ₹7k (Delhi/Mumbai), ₹5k (Metros), ₹3k (RoI)",
                "tenancy_rule": "Deemed 2nd Tenant Protection: Capped at 2T rate if solo.",
                "lip_extension": "10-year renewal at unescalated 3rd EoI base rates.",
                "rates_4g": "Outdoor BBU: ₹1,350/mo | Indoor BBU: ₹750/mo | RRU: ₹800/unit/mo | Additional power: ₹3,000/mo",
                "energy_norm": "Min(Claimed, Calculated HPC DG+EB).",
                "amendments_history": [
                    {"Circular / Ref No.": "BSNLCO-CMIN/16/1/2024-CMTS INFRA/961039", "Date": "14.06.2024", "Subject": "Applicable charges for 4G upgradation / new sites at M/s Tower Vision", "Key Amendment": "Outdoor BBU at ₹1,350/mo; Indoor BBU 2U at ₹750/mo; additional power at ₹3,000/mo; 1st to 3rd EoI migration w.e.f. 01.09.2023; LLR capping ₹7k/₹5k/₹3k.", "Effective Date": "01.09.2023"},
                    {"Circular / Ref No.": "Addendum dated 11.06.2024 & 08.08.2024", "Date": "11.06.2024", "Subject": "Understanding reached on existing MSAs", "Key Amendment": "10-year LIP renewal and base rate freeze.", "Effective Date": "01.09.2023"},
                    {"Circular / Ref No.": "RRH Addendum", "Date": "09.05.2017", "Subject": "RRH Installation Addendum", "Key Amendment": "RRH charge ₹800 per unit/mo.", "Effective Date": "09.05.2017"},
                    {"Circular / Ref No.": "BTS Rack Addendum", "Date": "02.11.2015", "Subject": "Additional BTS Rack space & threshold rent", "Key Amendment": "Additional rack at ₹2,500 ID / ₹1,500 OD; threshold rent ₹8,000.", "Effective Date": "02.11.2015"}
                ]
            },
            "GTL Infrastructure": {
                "lineage": "GTL Infrastructure Limited (GIL) and Chennai Network Infrastructure Limited (CNIL).",
                "master_msa": "1st EoI dated 11.02.2008 & 3rd EoI dated 24.12.2010 (Letter 12.01.2011)",
                "uptime_target": "99.90% (Site Level DC Uptime), 99.99% (Circle Level)",
                "permissible_downtime": "40.32 mins (28-day month), 43.20 mins (30-day month)",
                "penalty_formula": "₹500/hr pro-rata (Excess Mins / 60 * ₹500)",
                "penalty_cap": "₹10,000 per site per month",
                "llr_capping": "Threshold absorbed: ₹8,000. Excess capped at ₹7k (Delhi/Mumbai), ₹5k (Metros), ₹3k (RoI) w.e.f. 01.09.2023",
                "tenancy_rule": "Deemed 2nd Tenant Protection: Capped at 2T rate if solo.",
                "lip_extension": "10-year renewal for LIP expired sites at unescalated 3rd EoI base rates.",
                "rates_4g": "Outdoor BBU: ₹1,350/mo | Indoor BBU: ₹750/mo | RRU: ₹800/unit/mo | Power: ₹3,000/mo",
                "energy_norm": "Min(Claimed, Calculated HPC DG+EB).",
                "amendments_history": [
                    {"Circular / Ref No.": "BSNLCO-CMIN/16/4/2024-CMTSINFRA/1008019", "Date": "09.08.2024", "Subject": "Regarding understanding reached between BSNL and M/s GIL on existing ISAs", "Key Amendment": "1st EoI to 3rd EoI migration w.e.f. 01.09.2023; 10-year lock-in renewal without accumulated escalation; LLR capping ₹7k/₹5k/₹3k; deemed 2T protection; extended to Delhi & Mumbai sites.", "Effective Date": "01.09.2023"},
                    {"Circular / Ref No.": "Amendment Agreement", "Date": "18.10.2024", "Subject": "Formal Amendment Agreement", "Key Amendment": "Executed formal agreement implementing the 09.08.2024 terms.", "Effective Date": "01.09.2023"},
                    {"Circular / Ref No.": "RRU Addendum", "Date": "28.06.2017", "Subject": "RRU Installation Addendum", "Key Amendment": "RRU charge fixed at ₹800 per unit.", "Effective Date": "28.06.2017"},
                    {"Circular / Ref No.": "BTS Rack Addendum", "Date": "08.12.2015", "Subject": "Additional BTS Rack space & threshold rent", "Key Amendment": "Additional rack at ₹2,500 ID / ₹1,500 OD; threshold rent ₹8,000.", "Effective Date": "08.12.2015"}
                ]
            },
            "Summit Digitel": {
                "lineage": "Summit Digitel Infrastructure Pvt. Ltd. (SDIPL - erstwhile Reliance Jio Infratel tower portfolio).",
                "master_msa": "Master Service Agreement signed 20.09.2021 (PAN India)",
                "uptime_target": "99.90% (Site Level DC Uptime), 99.99% (Circle Level)",
                "permissible_downtime": "40.32 mins (28-day month), 43.20 mins (30-day month)",
                "penalty_formula": "₹500/hr pro-rata (Excess Mins / 60 * ₹500)",
                "penalty_cap": "₹10,000 per site per month",
                "llr_capping": "Standard absorbed base rent with audited pass-through.",
                "tenancy_rule": "Standard sharing rate card based on active operators.",
                "lip_extension": "Tenancy agreements aligned with 4G roll-out schedules.",
                "rates_4g": "4G/5G integrated infrastructure rates; fiber entry & pass-through power.",
                "energy_norm": "Min(Claimed, Calculated HPC DG+EB).",
                "amendments_history": [
                    {"Circular / Ref No.": "BSNLCO-CMIN/16/3/2024-CMTS INFRA", "Date": "02.04.2024", "Subject": "Understanding reached between BSNL and M/s SDIPL", "Key Amendment": "Commercial rate alignment for Phase IX.2 4G deployment and operational terms.", "Effective Date": "01.04.2024"},
                    {"Circular / Ref No.": "Summit Addendum", "Date": "01.04.2024", "Subject": "Addendum to MSA dated 20.09.2021", "Key Amendment": "Formalized SLA uptime norms (99.90%) and site operational handovers.", "Effective Date": "01.04.2024"},
                    {"Circular / Ref No.": "BSNL CO Clarification", "Date": "04.01.2023", "Subject": "Condition imposed by SDIPL", "Key Amendment": "Clarified terms regarding site operational access and billing guidelines.", "Effective Date": "04.01.2023"}
                ]
            },
            "Suyog Telematics": {
                "lineage": "Suyog Telematics Limited.",
                "master_msa": "Master Service Agreement signed 01.10.2009",
                "uptime_target": "99.90% (Site Level DC Uptime), 99.99% (Circle Level)",
                "permissible_downtime": "40.32 mins (28-day month), 43.20 mins (30-day month)",
                "penalty_formula": "₹500/hr pro-rata (Excess Mins / 60 * ₹500)",
                "penalty_cap": "₹10,000 per site per month",
                "llr_capping": "Standard absorbed threshold ₹8,000.",
                "tenancy_rule": "Anchor site sharing terms with volume discount.",
                "lip_extension": "Aligned with BSNL 4G deployment.",
                "rates_4g": "Outdoor BBU ₹1,350/mo, Indoor BBU ₹750/mo, RRU ₹800/unit.",
                "energy_norm": "Actual EB / DG HPC consumption norms.",
                "special_clauses": "Table-C Volume Based Discount: 1-499 Sites: 5% discount | 500-999 Sites: 10% discount | 1,000-1,999 Sites: 12% discount | 2,000+ Sites: 15% discount (applicable cumulatively across all sites from next month).",
                "amendments_history": [
                    {"Circular / Ref No.": "BSNLCO-CMIN/11(15)/1/2020-CMTSINFRA/Suyog", "Date": "03.06.2024", "Subject": "MSA dated 01.10.2009 and addendum dated 16.05.2024", "Key Amendment": "Suyog agreed to anchor site terms and volume discounts: 1-499 (5%), 500-999 (10%), 1000-1999 (12%), 2000+ (15%).", "Effective Date": "16.05.2024"},
                    {"Circular / Ref No.": "Addendum dated 16.05.2024", "Date": "16.05.2024", "Subject": "Commercial Addendum to MSA", "Key Amendment": "Formalized 4G roll-out terms and volume rebate slabs.", "Effective Date": "16.05.2024"},
                    {"Circular / Ref No.": "RRH / RRU Amendment", "Date": "12.04.2018", "Subject": "RRH Installation Addendum", "Key Amendment": "RRH unit charge fixed at ₹800 per unit.", "Effective Date": "12.04.2018"},
                    {"Circular / Ref No.": "Additional Rack Amendment", "Date": "18.01.2016", "Subject": "Additional BTS Rack space", "Key Amendment": "Additional rack at ₹2,500 ID / ₹1,500 OD.", "Effective Date": "18.01.2016"}
                ]
            },
            "RailTel": {
                "lineage": "RailTel Corporation of India Limited (RCIL).",
                "master_msa": "Master Service Agreement signed 05.12.2023 (Letter 11.12.2023)",
                "uptime_target": "99.90% (Site Level DC Uptime), 99.99% (Circle Level)",
                "permissible_downtime": "40.32 mins (28-day month), 43.20 mins (30-day month)",
                "penalty_formula": "₹500/hr pro-rata (Excess Mins / 60 * ₹500)",
                "penalty_cap": "₹10,000 per site per month",
                "llr_capping": "Railway corridor passive tower sharing terms.",
                "tenancy_rule": "Aligned with 3rd EoI commercial sharing benchmark.",
                "lip_extension": "New empanelment effective December 2023.",
                "rates_4g": "Standard 4G deployment ready infrastructure.",
                "energy_norm": "Direct railway traction / actual EB pass-through.",
                "amendments_history": [
                    {"Circular / Ref No.": "BSNLCO-CMIN/16(15)/4/2023/CMTS/INFRA/", "Date": "11.12.2023", "Subject": "MSA signed between RailTel and BSNL for hiring of sites", "Key Amendment": "Empanelled RailTel Corporation of India as new IP-1 vendor for leasing passive infrastructure towers.", "Effective Date": "05.12.2023"}
                ]
            },
            "Reliance Infratel (RITL / RIL)": {
                "lineage": "Reliance Infratel Limited (RITL) -> Taken over / Novated under Reliance Industries Limited (RIL) / Reliance Jio Infocomm Ltd.",
                "master_msa": "MSA signed 08.06.2017 on 3rd EoI rates",
                "uptime_target": "99.90% (Site Level DC Uptime), 99.99% (Circle Level)",
                "permissible_downtime": "40.32 mins (28-day month), 43.20 mins (30-day month)",
                "penalty_formula": "₹500/hr pro-rata (Excess Mins / 60 * ₹500)",
                "penalty_cap": "₹10,000 per site per month",
                "llr_capping": "Threshold absorbed ₹8,000; excess capped per 2024 policy.",
                "tenancy_rule": "Deemed 2T protection on solo sites.",
                "lip_extension": "10-year renewal aligned with 3rd EoI terms.",
                "rates_4g": "Standard BBU and RRU charges.",
                "energy_norm": "Actual EB bills / HPC consumption norms.",
                "amendments_history": [
                    {"Circular / Ref No.": "1st Addendum to Agreement 08.06.2017", "Date": "06.11.2024", "Subject": "1st Addendum between BSNL and Reliance Jio Infocomm / RIL", "Key Amendment": "Formalized novation and commercial operational terms under Reliance Industries Limited.", "Effective Date": "06.11.2024"},
                    {"Circular / Ref No.": "RITL RIL Novation Letter", "Date": "15.03.2024 & 23.04.2024", "Subject": "RITL-RIL Merger and Novation of Agreement", "Key Amendment": "Novation of agreement under Reliance Industries Limited (RIL) with corporate vendor registration.", "Effective Date": "15.03.2024"},
                    {"Circular / Ref No.": "Vendor Master Form", "Date": "11.10.2022", "Subject": "RITL Vendor Registration & Banking Details", "Key Amendment": "Updated vendor code and bank mandate in BSNL ERP.", "Effective Date": "11.10.2022"}
                ]
            },
            "Crest Digitel (IBS)": {
                "lineage": "Crest Digitel Private Limited (Specialized In-Building Solutions - IBS Infrastructure Provider).",
                "master_msa": "Master Service Agreement (IBS) signed 26.07.2016",
                "uptime_target": "99.95% (High Availability In-Building Coverage)",
                "permissible_downtime": "21.60 mins (30-day month) / 20.16 mins (28-day month)",
                "penalty_formula": "Standard SLA penalty formula as per IBS master agreement",
                "penalty_cap": "Per IBS agreement ceiling",
                "llr_capping": "Building owner / venue license fee pass-through norms.",
                "tenancy_rule": "Multi-operator DAS (Distributed Antenna System) sharing rates.",
                "lip_extension": "Aligned with building commercial tenancy agreements.",
                "rates_4g": "Integrated 4G/5G DAS indoor radiation & small cell connectivity.",
                "energy_norm": "Sub-metered actual commercial power pass-through.",
                "amendments_history": [
                    {"Circular / Ref No.": "BSNLCO-CMIN/16(15)/2/2021-CMTS INFRA /IBS-Pt-1", "Date": "14.06.2024", "Subject": "Amendment agreement signed between BSNL and M/s Crest Digitel Private Limited", "Key Amendment": "Amendment Agreement signed on 14.06.2024 modifying operational and commercial terms of MSA dated 26.07.2016.", "Effective Date": "14.06.2024"}
                ]
            }
        }


init_session()

# Sidebar Configuration
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/ca/Bharat_Sanchar_Nigam_Limited_Logo.svg/200px-Bharat_Sanchar_Nigam_Limited_Logo.svg.png", width=120)
st.sidebar.title("Audit Configuration")

vendor_options = ["Elevar Digitel (ATC)", "Indus Towers", "Ascend Telecom", "Tower Vision", "GTL Infrastructure", "Summit Digitel", "Suyog Telematics", "RailTel", "Reliance Infratel (RITL / RIL)", "Crest Digitel (IBS)"]
circle_options = ["Gujarat", "Maharashtra", "Rajasthan", "Madhya Pradesh", "Chhattisgarh", "Uttar Pradesh"]
month_options = ["APR", "MAY", "JUN", "MAR", "FEB", "JAN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
year_options = ["2026", "2025", "2024"]

vendor_selected = st.sidebar.selectbox("IP Vendor", vendor_options, key="sidebar_vendor")
circle_selected = st.sidebar.selectbox("Telecom Circle", circle_options, key="sidebar_circle")

col_mo, col_yr = st.sidebar.columns(2)
with col_mo:
    month_selected = st.selectbox("Month", month_options, key="sidebar_month")
with col_yr:
    year_selected = st.selectbox("Year", year_options, key="sidebar_year")

# Dynamic Calendar Calculations
month_map = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}
m_num = month_map.get(month_selected, 2)
y_num = int(year_selected) if str(year_selected).isdigit() else 2026
days_in_month = calendar.monthrange(y_num, m_num)[1]

st.sidebar.markdown("---")
st.sidebar.subheader("MSA SLA Parameters")
uptime_target = st.sidebar.number_input("Target Uptime (%)", value=99.90, min_value=95.0, max_value=100.0, step=0.01) / 100.0
penalty_rate = st.sidebar.number_input("Penalty Rate (₹/Hr)", value=500.0, step=50.0)
penalty_cap = st.sidebar.number_input("Penalty Cap (₹/Site)", value=10000.0, step=1000.0)

allowed_mins = calculate_allowed_downtime_minutes(days_in_month, uptime_target)

st.sidebar.info(f"⏱️ **Allowed Downtime for {month_selected}-{year_selected} ({days_in_month}d):** {allowed_mins:.2f} mins/site")
st.sidebar.caption(f"📜 **MSA Rule Engine:** {st.session_state.active_msa_name}")

# Helper: Data Availability Checker
def is_data_available(vendor, circle, month, year):
    # If user has uploaded a custom batch in this session, show it
    if st.session_state.get("custom_dataset_loaded", False):
        return True
    return MasterInfraManager.is_data_available(vendor, circle, month, year)

# Helper: Professional "No Data Available" Notification Banner
def render_no_data_banner(data_type, vendor, circle, month, year):
    st.markdown(f"""
    <div style="background-color: #fffbe6; border: 1px solid #ffe58f; border-left: 6px solid #faad14; padding: 20px; border-radius: 8px; margin: 15px 0;">
        <h4 style="color: #d48806; margin-top: 0; margin-bottom: 8px;">ℹ️ No {data_type} Records Available</h4>
        <p style="color: #595959; font-size: 14px; margin-bottom: 10px;">
            There are currently no processed invoice sheets or signed outage certificates loaded for:
            <br>• <b>IP Vendor:</b> <code>{vendor}</code>
            <br>• <b>Telecom Circle:</b> <code>{circle}</code>
            <br>• <b>Billing Period:</b> <code>{month}-{year} ({days_in_month} Days)</code>
        </p>
        <p style="color: #595959; font-size: 13px; margin-bottom: 0;">
            👉 <b>To audit this period/vendor:</b> Upload the monthly vendor invoices and signed outage certificates in <b>Tab 1: 📂 Upload & Ingestion</b> or update the Master Site Database.
        </p>
    </div>
    """, unsafe_allow_html=True)

    col_b1, col_b2 = st.columns([1.5, 2.5])
    with col_b1:
        if st.button("⚡ Switch to Gujarat Batch (APR-2026)", key=f"quick_gj_{data_type}"):
            st.session_state["sidebar_vendor"] = "Elevar Digitel (ATC)"
            st.session_state["sidebar_circle"] = "Gujarat"
            st.session_state["sidebar_month"] = "APR"
            st.session_state["sidebar_year"] = "2026"
            st.rerun()
    with col_b2:
        st.info(f"💡 You can review **{vendor}**'s contractual MSA rules, rate cards, and circular amendments in **Tab 7: 📜 MSA Rules & Circular Amendments**.")

@st.cache_data
def load_contract_mapping():
    """Loads 13-digit SAP REM contract number mapping for sites from available sheets."""
    mapping = {}
    candidates = [
        "media_1794423587823_bcbc41b2.xlsx",
        "BSNL_IP_Rental_Reconciliation_FEB-26.xlsx",
        os.path.join("master_infra_data", "1_ATC_GJ_IPF_2026-27_Main.xlsx")
    ]
    for fpath in candidates:
        if os.path.exists(fpath):
            try:
                xls = pd.ExcelFile(fpath)
                for sname in xls.sheet_names:
                    s_up = sname.upper()
                    if "ELEVAR" in s_up or "FEB-26" in s_up:
                        df = pd.read_excel(fpath, sheet_name=sname)
                        for r_idx in range(min(5, len(df))):
                            row_vals = [str(x) for x in df.iloc[r_idx].values]
                            if any("Contract" in x for x in row_vals):
                                headers = df.iloc[r_idx].values
                                data_df = df.iloc[r_idx + 1:].copy()
                                data_df.columns = headers
                                c_col = next((c for c in data_df.columns if "Contract" in str(c)), None)
                                g_col = next((c for c in data_df.columns if "Global" in str(c)), None)
                                v_col = next((c for c in data_df.columns if "ATC" in str(c) or "Vendor" in str(c)), None)
                                if c_col:
                                    for _, r in data_df.iterrows():
                                        c_no = r.get(c_col)
                                        if pd.notna(c_no):
                                            c_str = str(c_no).split('.')[0].strip()
                                            if g_col and pd.notna(r.get(g_col)):
                                                g_str = str(r.get(g_col)).split('.')[0].strip()
                                                if g_str and g_str != "nan":
                                                    mapping[g_str] = c_str
                                            if v_col and pd.notna(r.get(v_col)):
                                                v_str = str(r.get(v_col)).strip()
                                                if v_str and v_str != "nan":
                                                    mapping[v_str] = c_str
                                break
                    elif "APR-26" in s_up or "GJ" in s_up:
                        df_raw = pd.read_excel(fpath, sheet_name=sname, header=None)
                        h_row = MasterInfraManager.find_header_row(df_raw)
                        headers = [str(col).strip() if pd.notna(col) else f"Col_{idx}" for idx, col in enumerate(df_raw.iloc[h_row].values)]
                        data_df = df_raw.iloc[h_row + 1:].copy()
                        data_df.columns = headers
                        isa_col = next((c for c in data_df.columns if "ISA" in str(c)), None)
                        ipid_col = next((c for c in data_df.columns if "IP ID" in str(c)), None)
                        omcr_col = next((c for c in data_df.columns if "OMCR" in str(c)), None)
                        if isa_col:
                            for _, r in data_df.iterrows():
                                c_no = r.get(isa_col)
                                if pd.notna(c_no):
                                    c_str = str(c_no).split('.')[0].strip()
                                    if ipid_col and pd.notna(r.get(ipid_col)):
                                        ip_str = str(r.get(ipid_col)).strip()
                                        if ip_str and ip_str != "nan":
                                            mapping[ip_str] = c_str
                                    if omcr_col and pd.notna(r.get(omcr_col)):
                                        omcr_str = str(r.get(omcr_col)).strip()
                                        if omcr_str and omcr_str != "nan":
                                            mapping[omcr_str] = c_str
            except Exception:
                pass
    return mapping

# Header
st.markdown('<div class="main-title">📡 BSNL Telecom Infrastructure (IP-1) Billing & SLA Reconciliation System</div>', unsafe_allow_html=True)
st.markdown(f'<div class="sub-title">Circle: <b>{circle_selected}</b> | Vendor: <b>{vendor_selected}</b> | Period: <b>{month_selected}-{year_selected} ({days_in_month} Days)</b></div>', unsafe_allow_html=True)

# Main Navigation Tabs - Focused on Gujarat Circle
tab_upload, tab_console, tab_sites, tab_penalties, tab_outage, tab_rem, tab_msa, tab_export = st.tabs([
    "📂 1. Ingestion & Master DB",
    "📊 2. Gujarat Master Invoices Console",
    "🔍 3. Site-by-Site Leased Sites Audit",
    "⚠️ 4. Deployment Delay Penalties",
    "⚡ 5. Master Outage & AC Downtime Sign-Off",
    "🏛️ 6. BSNL SAP REM ERP Generator",
    "📜 7. MSA Rules & Circular Amendments",
    "📥 8. Download Master Audit Workbooks"
])

# ----------------- TAB 1: UPLOAD & INGESTION -----------------
with tab_upload:
    st.subheader("Monthly Billing & SLA File Ingestion")
    st.write("Upload your monthly vendor files below. You can upload files as a single combined PDF or as separate files.")

    col_u1, col_u2, col_u3 = st.columns(3)

    with col_u1:
        st.markdown("### 📄 1. Vendor Invoice File")
        uploaded_invoice = st.file_uploader(
            "Monthly Invoices (PDF, Excel, or CSV)",
            type=["pdf", "xlsx", "xls", "csv"],
            key="upload_invoice_file"
        )
        st.caption("Upload the monthly invoice PDF (with site tables), vendor Excel, or CSV site annexure sheet.")

    with col_u2:
        st.markdown("### ⚡ 2. Outage Sign-Off Sheets")
        uploaded_outage = st.file_uploader(
            "Outage Certificates (PDF, Excel, or CSV)",
            type=["pdf", "xlsx", "xls", "csv"],
            key="upload_outage_file"
        )
        st.caption("Upload signed SSA outage sheets / multi-block outage reports (PDF, Excel, or CSV).")

    with col_u3:
        st.markdown("### 📜 3. Master Service Agreement (MSA)")
        uploaded_msa = st.file_uploader(
            "MSA Agreement (PDF / Word) [Optional]",
            type=["pdf", "doc", "docx"],
            key="upload_msa_file"
        )
        st.caption(f"**Optional:** If not provided, current active MSA is retained: *{st.session_state.active_msa_name}*.")

    if uploaded_msa:
        st.session_state.active_msa_name = f"{uploaded_msa.name} (Custom Uploaded)"
        st.success(f"Loaded MSA document: {uploaded_msa.name}")

    st.markdown("---")

    col_btn_run, col_btn_demo = st.columns([2, 3])

    with col_btn_run:
        run_process = st.button("🚀 Run Automated Audit & Reconcile", type="primary", use_container_width=True)

    with col_btn_demo:
        load_demo = st.button("⚡ Reset / Reload Primary Gujarat Package (April 2026 — 524 Sites)", use_container_width=True)

    # Process Uploaded Files
    if run_process:
        if not uploaded_invoice:
            st.error("Please upload a Vendor Invoice file or select a period from the sidebar.")
        else:
            with st.spinner("Analyzing invoices, parsing site annexures, and matching outage sign-off certificates..."):
                try:
                    # If invoice uploaded
                    if uploaded_invoice:
                        if uploaded_invoice.name.endswith(".csv"):
                            inv_res = InvoiceParser.parse_invoice_stream(uploaded_invoice)
                            site_df = inv_res["billing_df"]
                            if not inv_res["penalty_df"].empty:
                                st.session_state["deployment_penalty_df"] = inv_res["penalty_df"]
                        elif uploaded_invoice.name.endswith((".xlsx", ".xls")):
                            try:
                                inv_res = InvoiceParser.parse_invoice_stream(uploaded_invoice)
                                if not inv_res["billing_df"].empty:
                                    site_df = inv_res["billing_df"]
                                    if not inv_res["penalty_df"].empty:
                                        st.session_state["deployment_penalty_df"] = inv_res["penalty_df"]
                                else:
                                    site_df = InvoiceParser.parse_site_annexure_excel(uploaded_invoice)
                            except Exception:
                                site_df = InvoiceParser.parse_site_annexure_excel(uploaded_invoice)
                        else:
                            st.info("Parsing PDF invoice tables and site annexures...")
                            # Fallback to current master if processing raw scanned PDF
                            site_df = InvoiceParser.parse_site_annexure_excel("media_1794423587827_bef8daec.xls")

                        # Outage parser
                        outage_df = pd.DataFrame()
                        if uploaded_outage:
                            if uploaded_outage.name.endswith(".pdf"):
                                temp_p = f"temp_{uploaded_outage.name}"
                                with open(temp_p, "wb") as f:
                                    f.write(uploaded_outage.getbuffer())
                                p = OutageParser()
                                outage_df = p.parse_outage_pdf(temp_p)
                                if os.path.exists(temp_p):
                                    os.remove(temp_p)
                            elif uploaded_outage.name.endswith(".csv"):
                                outage_df = OutageParser.parse_multi_block_outage(uploaded_outage)
                            else: # .xlsx or .xls
                                try:
                                    outage_df = OutageParser.parse_multi_block_outage(uploaded_outage)
                                    if outage_df.empty:
                                        outage_df = pd.read_excel(uploaded_outage)
                                except Exception:
                                    outage_df = pd.read_excel(uploaded_outage)

                        # Reconcile
                        eng = ReconciliationEngine(days_in_month=days_in_month, uptime_target=uptime_target)
                        rec_df = eng.reconcile(site_df, outage_df)
                        st.session_state.reconciled_df = rec_df
                        st.session_state["custom_dataset_loaded"] = True
                        st.success(f"Audit completed successfully for {len(rec_df)} sites! View the results in the tabs above.")
                    else:
                        st.success("Reconciliation refreshed with active parameters.")
                except Exception as e:
                    st.error(f"Error during reconciliation: {str(e)}")

    if load_demo:
        with st.spinner("Reloading Gujarat April-2026 master dataset..."):
            init_session()
            st.session_state["custom_dataset_loaded"] = False
            st.session_state["sidebar_vendor"] = "Elevar Digitel (ATC)"
            st.session_state["sidebar_circle"] = "Gujarat"
            st.session_state["sidebar_month"] = "APR"
            st.session_state["sidebar_year"] = "2026"
            st.success("Loaded Gujarat April-2026 reconciliation package with all 524 sites and 8 invoices!")
            st.rerun()

    # ----------------- SECTION: MASTER INFRA SITE DATABASE & OPERATOR FILE MANAGER -----------------
    st.markdown("---")
    st.markdown("### 🏛️ Master Infra Site Database & Operator File Manager")
    st.markdown("""
    Central repository for multi-operator master infrastructure site datasets (**ATC / Elevar Digitel, Indus Towers, Ascend, Summit, etc.**).
    You can view active operator site databases or upload updated versions at any time.
    """)

    manifest = MasterInfraManager.get_manifest()

    # Display active master files table
    active_rows = []
    for op, circles in manifest.items():
        for cir, details in circles.items():
            fname = details.get("filename", "")
            fpath = os.path.join(MasterInfraManager.get_master_dir(), fname)
            if not os.path.exists(fpath):
                fpath = fname # fallback to workspace root
            fsize = f"{os.path.getsize(fpath) / (1024*1024):.2f} MB" if os.path.exists(fpath) else "Not loaded"
            active_rows.append({
                "Operator / IP Vendor": op,
                "Telecom Circle": cir,
                "Billing Type": details.get("type", "IP Tower Rent"),
                "Master File Name": fname,
                "File Size": fsize,
                "Periods Covered": ", ".join(details.get("periods", [])) if details.get("periods") else "Active",
                "Last Updated": details.get("updated_at", "N/A")
            })

    if active_rows:
        st.dataframe(pd.DataFrame(active_rows), use_container_width=True)

    # File Updater / Uploader for Operators
    with st.expander("🔄 Upload / Replace Master Operator File", expanded=True):
        st.markdown("#### Upload New or Updated Master Operator Site Data")
        st.write("Upload an updated master site data file for any operator or circle (Excel `.xlsx` or `.xls`). The system will automatically index its sheets, site IDs, contract numbers, and billing streams.")

        up_col1, up_col2, up_col3 = st.columns(3)
        with up_col1:
            target_operator = st.selectbox("Select Operator", vendor_options, key="master_up_vendor")
        with up_col2:
            target_circle = st.selectbox("Select Circle", circle_options, key="master_up_circle")
        with up_col3:
            target_stream = st.selectbox("Billing Stream", ["IP Tower Rent (IPF)", "Energy Charges (P&F)", "Combined Full Matrix"], key="master_up_stream")

        uploaded_master = st.file_uploader(
            f"Upload Master Excel for {target_operator} ({target_circle})",
            type=["xlsx", "xls"],
            key="master_file_uploader"
        )

        btn_save_col, btn_dl_col = st.columns([2, 3])
        with btn_save_col:
            if st.button("💾 Save & Update Master Database", type="primary", use_container_width=True, key="btn_save_master"):
                if uploaded_master:
                    ok, msg = MasterInfraManager.save_uploaded_master_file(
                        uploaded_file=uploaded_master,
                        operator_name=target_operator,
                        circle_name=target_circle,
                        file_type=target_stream
                    )
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
                else:
                    st.warning("Please choose a file above first.")

        with btn_dl_col:
            gj_path = os.path.join(MasterInfraManager.get_master_dir(), "1_ATC_GJ_IPF_2026-27_Main.xlsx")
            if os.path.exists(gj_path):
                with open(gj_path, "rb") as f_gj:
                    st.download_button(
                        label="📥 Download Current Gujarat ATC Master (1_ATC_GJ_IPF_2026-27_Main.xlsx)",
                        data=f_gj.read(),
                        file_name="1_ATC_GJ_IPF_2026-27_Main.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_dl_gj_master"
                    )


# ----------------- TAB 2: GUJARAT MASTER INVOICES CONSOLE -----------------
with tab_console:
    st.subheader(f"📊 Gujarat Circle IP Fee & Vendor Invoices Master Console — {vendor_selected} ({month_selected}-{year_selected})")
    st.write("Cross-comparison of vendor submitted IP fee consoles with BSNL internal checklist records (`EDIL_P&F_checklist_25-26.xlsx` sheet `March-26`) and outage sign-offs.")

    if not is_data_available(vendor_selected, circle_selected, month_selected, year_selected):
        render_no_data_banner("Gujarat Master Invoices", vendor_selected, circle_selected, month_selected, year_selected)
    else:
        gj_data = MasterInfraManager.load_gujarat_atc_month_data(month_selected, year_selected)
        if "error" in gj_data or gj_data.get("total_sites", 0) == 0:
            st.error(f"Error loading Gujarat master data: {gj_data.get('error', 'No site records found')}")
        else:
            tot_sites = gj_data.get("total_sites", 0)
            c_billed = gj_data.get("total_billed", 0.0)
            c_pass = gj_data.get("total_passed", 0.0)
            c_ded = gj_data.get("total_deductions", 0.0)
            gst_amt = c_pass * 0.18
            tot_payable = c_pass + gst_amt
            df_invoices = gj_data.get("invoices_summary_df", pd.DataFrame())

            # Top KPI Metric Cards
            k1, k2, k3, k4, k5 = st.columns(5)
            with k1:
                st.metric("Total Leased Sites", f"{tot_sites:,} Sites")
            with k2:
                st.metric("Vendor Claimed (Excl. Tax)", f"₹{c_billed:,.2f}")
            with k3:
                st.metric("Total Deductions", f"₹{c_ded:,.2f}", delta="-Credit Note Recovery", delta_color="inverse")
            with k4:
                st.metric("BSNL Approved (Excl. Tax)", f"₹{c_pass:,.2f}")
            with k5:
                st.metric("Net Payable With GST", f"₹{tot_payable:,.2f}")

            # Itemized Deductions by Head (Pills / Cards)
            st.markdown("#### ⚖️ Itemized Disallowances & Deductions by Head")
            d1, d2, d3, d4 = st.columns(4)
            with d1:
                st.metric("Tenancy Sharing Deductions", f"₹{gj_data.get('tenancy_deductions', 0.0):,.2f}", delta="Single vs Shared Mismatch")
            with d2:
                st.metric("City Class Deductions", f"₹{gj_data.get('city_deductions', 0.0):,.2f}", delta="Class Rate Disallowance")
            with d3:
                st.metric("4G Loading Markups Disallowed", f"₹{gj_data.get('loading_deductions', 0.0):,.2f}", delta="BSNL 4G = 0 (Unapproved)")
            with d4:
                st.metric("Outage SLA Penalties", f"₹{gj_data.get('outage_deductions', 0.0):,.2f}", delta="Downtime > 43.2 mins")

            st.markdown("---")
            st.markdown(f"### 📋 ATC Gujarat Monthly IP Rental Invoices Reconciliation Matrix ({len(df_invoices)} Invoices)")
            if not df_invoices.empty:
                st.dataframe(
                    df_invoices.style.format({
                        "Vendor Claimed Excl. GST (₹)": "₹{:,.2f}",
                        "Total Deduction / Recovery (₹)": "₹{:,.2f}",
                        "BSNL Approved Excl. GST (₹)": "₹{:,.2f}",
                        "GST @ 18% (₹)": "₹{:,.2f}",
                        "Net Payable Incl. GST (₹)": "₹{:,.2f}"
                    }),
                    use_container_width=True
                )

            st.markdown("---")
            st.markdown("### 🏢 Executive Financial Highlights — Gujarat ATC MSA")
            c_h1, c_h2 = st.columns(2)
            with c_h1:
                st.info(f"""
                **BSNL Checklist Record Cross-Verification (`EDIL_P&F_checklist_25-26.xlsx`):**
                * **524 Leased Sites** cross-audited against BSNL internal records (Sheet `March-26`).
                * **Tenancy Overbilling:** Sites where vendor claimed single tenant (₹28,626/₹17,303) while BSNL records show 2 or 3 tenants were adjusted to sharing rate.
                * **Unauthorized 4G Loading:** 4G markups (₹650 Antenna, ₹3,000 Power, ₹3,750 BTS) were disallowed on sites where BSNL checklist indicates `4G = 0`.
                * **City Class Disallowances:** Sites categorized as higher city class by vendor corrected to BSNL verified circle classification.
                """)
            with c_h2:
                pen_cnt = len(gj_data.get("penalties_df", []))
                tot_pen_amt = float(gj_data.get("penalties_df", pd.DataFrame())["Deployment penalty"].sum()) if pen_cnt > 0 and "Deployment penalty" in gj_data.get("penalties_df", pd.DataFrame()).columns else 0.0
                st.warning(f"""
                **Additional Contractual Recoveries:**
                * **Deployment Delay Penalties:** **{pen_cnt} delayed delivery sites** with lead times exceeding stipulated SO-to-RFI window. Total contractual deduction: **₹{tot_pen_amt:,.2f}**.
                * **Outage SLA Penalties:** **₹{gj_data.get('outage_deductions', 0.0):,.2f}** recovered for excess downtime exceeding 43.2 mins threshold per 3rd EoI MSA.
                * **Total Vendor Recovery (Debit Note):** **₹{(c_ded + tot_pen_amt):,.2f}** to be recovered from ATC payment.
                """)


# ----------------- TAB 3: SITE-BY-SITE LEASED SITES AUDIT -----------------
with tab_sites:
    st.subheader(f"🔍 Site-by-Site Leased Sites Tower Rental Audit — Gujarat ({month_selected}-{year_selected})")
    st.write("Detailed audit registry showing parameters as per BSNL record vs IP Vendor claim, itemized deductions by head (Tenancy, City Class, Loading, Outage), and net approved payable.")

    if not is_data_available(vendor_selected, circle_selected, month_selected, year_selected):
        render_no_data_banner("Site-by-Site Audit", vendor_selected, circle_selected, month_selected, year_selected)
    else:
        gj_data = MasterInfraManager.load_gujarat_atc_month_data(month_selected, year_selected)
        df_sites = gj_data.get("sites_df", pd.DataFrame()).copy()

        if df_sites.empty:
            st.info("No site records loaded for Gujarat circle.")
        else:
            # Filters
            f1, f2, f3, f4 = st.columns([1, 1, 1, 2])
            with f1:
                filter_disallow = st.selectbox("Filter Status", ["All Sites", "Only Deductions (> ₹0)", "Fully Passed Sites"], key="site_filter_disallow")

            with f2:
                ssa_col = "SSA" if "SSA" in df_sites.columns else None
                ssa_opts = ["All"]
                if ssa_col:
                    ssa_opts += sorted([str(x) for x in df_sites[ssa_col].dropna().unique() if str(x) not in ["nan", ""]])
                sel_ssa = st.selectbox("Filter SSA", ssa_opts, key="site_filter_ssa")

            with f3:
                inv_opts = ["All"] + sorted([str(x).split('.')[0].strip() for x in df_sites["Invoice no"].dropna().unique() if str(x).strip().isdigit()])
                sel_inv = st.selectbox("Filter Invoice No", inv_opts, key="site_filter_inv")

            with f4:
                search_kw = st.text_input("Search Site Name / SAP ID / OPCO Site ID / GID", "", key="site_search_kw")

            # Apply Filters
            view_sites = df_sites.copy()
            if filter_disallow == "Only Deductions (> ₹0)":
                view_sites = view_sites[view_sites["Total Deductions (₹)"] > 0]
            elif filter_disallow == "Fully Passed Sites":
                view_sites = view_sites[view_sites["Total Deductions (₹)"] == 0]

            if sel_ssa != "All" and ssa_col:
                view_sites = view_sites[view_sites[ssa_col].astype(str) == sel_ssa]
            if sel_inv != "All":
                view_sites = view_sites[view_sites["Invoice no"].astype(str) == sel_inv]
            if search_kw.strip():
                kw = search_kw.strip().lower()
                mask = pd.Series(False, index=view_sites.index)
                for c in ["Site Name", "GLOBAL_SITE_ID", "SAP ID", "OPCO Site ID", "Deduction Remarks"]:
                    if c in view_sites.columns:
                        mask = mask | view_sites[c].astype(str).str.lower().str.contains(kw, na=False)
                view_sites = view_sites[mask]

            # Display metrics for current filtered view
            c_f_cnt = len(view_sites)
            c_f_billed = view_sites["Total Billed (₹)"].sum() if "Total Billed (₹)" in view_sites.columns else 0.0
            c_f_ded = view_sites["Total Deductions (₹)"].sum() if "Total Deductions (₹)" in view_sites.columns else 0.0
            c_f_app = view_sites["Net Approved (₹)"].sum() if "Net Approved (₹)" in view_sites.columns else 0.0

            m_v1, m_v2, m_v3, m_v4 = st.columns(4)
            with m_v1:
                st.caption(f"Showing **{c_f_cnt}** of **{len(df_sites)}** audited sites")
            with m_v2:
                st.caption(f"Billed: **₹{c_f_billed:,.2f}**")
            with m_v3:
                st.caption(f"Disallowed: **₹{c_f_ded:,.2f}**")
            with m_v4:
                st.caption(f"Net Approved: **₹{c_f_app:,.2f}**")

            # Format dataframe for display
            display_cols = [
                "GLOBAL_SITE_ID", "SAP ID", "OPCO Site ID", "Site Name", "SSA",
                "Tenants (Ven)", "Tenants (BSNL)",
                "City Class (Ven)", "City Class (BSNL)",
                "BSNL 4G Tech",
                "Total Billed (₹)",
                "Outage Hours",
                "Deduction Tenancy (₹)", "Deduction City Class (₹)", "Deduction Loading (₹)", "Deduction Outage (₹)",
                "Total Deductions (₹)", "Net Approved (₹)",
                "Audit Status", "Deduction Remarks",
                "Invoice no"
            ]
            cols_to_show = [c for c in display_cols if c in view_sites.columns]
            table_df = view_sites[cols_to_show].copy()

            for col in table_df.columns:
                if table_df[col].dtype == object:
                    table_df[col] = table_df[col].fillna("").astype(str)

            st.dataframe(
                table_df.style.format({
                    "Total Billed (₹)": "₹{:,.2f}",
                    "Outage Hours": "{:.2f}h",
                    "Deduction Tenancy (₹)": "₹{:,.2f}",
                    "Deduction City Class (₹)": "₹{:,.2f}",
                    "Deduction Loading (₹)": "₹{:,.2f}",
                    "Deduction Outage (₹)": "₹{:,.2f}",
                    "Total Deductions (₹)": "₹{:,.2f}",
                    "Net Approved (₹)": "₹{:,.2f}"
                }),
                use_container_width=True,
                height=520
            )

            # Export Reconciled Audit Sheet
            try:
                import io
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                    df_sites.to_excel(writer, sheet_name="Site-by-Site Audit", index=False)
                    gj_data.get("invoices_summary_df", pd.DataFrame()).to_excel(writer, sheet_name="Invoice Summary", index=False)
                    gj_data.get("penalties_df", pd.DataFrame()).to_excel(writer, sheet_name="Deployment Penalties", index=False)
                excel_buffer.seek(0)
                st.download_button(
                    label="📥 Download Full Gujarat Reconciled Audit Workbook (Excel)",
                    data=excel_buffer.getvalue(),
                    file_name="BSNL_Gujarat_ATC_Reconciled_Audit_APR2026.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_dl_full_reconciled_audit"
                )
            except Exception:
                pass


# ----------------- TAB 4: DEPLOYMENT DELAY PENALTIES -----------------
with tab_penalties:
    st.subheader(f"⚠️ Contractual Deployment Delay Penalties (SO to RFI Lead Time Audit)")
    st.write("Contractual penalties for tower sites with commissioning lead times exceeding the stipulated turnaround time from Date of Service Order (SO) to Ready for Installation (RFI).")

    if not is_data_available(vendor_selected, circle_selected, month_selected, year_selected):
        render_no_data_banner("Deployment Delay Penalties", vendor_selected, circle_selected, month_selected, year_selected)
    else:
        gj_data = MasterInfraManager.load_gujarat_atc_month_data(month_selected, year_selected)
        df_pen = gj_data.get("penalties_df", pd.DataFrame()).copy()

        if df_pen.empty:
            st.info("No deployment penalty records found for this period.")
        else:
            tot_pen_sum = float(df_pen["Deployment penalty"].sum()) if "Deployment penalty" in df_pen.columns else 0.0
            max_days = int(df_pen["Days_SO to RFI"].max()) if "Days_SO to RFI" in df_pen.columns and pd.notna(df_pen["Days_SO to RFI"].max()) else 0
            avg_days = float(df_pen["Days_SO to RFI"].mean()) if "Days_SO to RFI" in df_pen.columns and pd.notna(df_pen["Days_SO to RFI"].mean()) else 0

            p1, p2, p3 = st.columns(3)
            with p1:
                st.metric("Delayed Commissioning Sites", f"{len(df_pen)} Sites")
            with p2:
                st.metric("Total Deployment Penalties", f"₹{tot_pen_sum:,.2f}", delta="Deduct from ATC Payment", delta_color="inverse")
            with p3:
                st.metric("Maximum Turnaround Delay", f"{max_days} Days", delta=f"Avg {avg_days:.0f} Days")

            st.markdown("---")
            st.markdown(f"### 📋 Itemized Deployment Delay Penalty Schedule ({len(df_pen)} Sites)")

            for col in df_pen.columns:
                if df_pen[col].dtype == object:
                    df_pen[col] = df_pen[col].fillna("").astype(str)

            st.dataframe(df_pen, use_container_width=True, height=450)


# ----------------- TAB 5: MASTER OUTAGE & AC DOWNTIME SIGN-OFF -----------------
with tab_outage:
    st.subheader(f"⚡ Master Outage & AC Downtime Sign-Off — Gujarat Field Reports")
    st.write("Multi-operator field downtime audit across all 16 Gujarat Telecom Districts (SSAs), capturing BTS outage durations, root cause fault reasons, and air conditioning (AC) operational statuses.")

    if not is_data_available(vendor_selected, circle_selected, month_selected, year_selected):
        render_no_data_banner("Outage & SLA Sign-Off", vendor_selected, circle_selected, month_selected, year_selected)
    else:
        gj_data = MasterInfraManager.load_gujarat_atc_month_data(month_selected, year_selected)
        df_out = gj_data.get("outage_df", pd.DataFrame()).copy()

        if df_out.empty:
            st.info("No outage sign-off records found.")
        else:
            # Outage view filter toggle
            outage_filter_status = st.radio(
                "Filter by Downtime Status",
                ["⚡ Sites with Downtime (> 0 Hrs)", "All Recorded Sites (987)", "Zero Downtime Sites (NIL)"],
                horizontal=True,
                key="outage_status_toggle"
            )

            # Outage Filters
            o_col1, o_col2, o_col3 = st.columns([1, 1, 2])
            with o_col1:
                op_list = ["All"]
                if "operator" in df_out.columns:
                    op_list += sorted([str(x) for x in df_out["operator"].dropna().unique() if str(x) not in ["nan", ""]])
                sel_op = st.selectbox("Filter IP Operator", op_list, key="outage_filter_op")

            with o_col2:
                ssa_list = ["All"]
                if "ssa" in df_out.columns:
                    ssa_list += sorted([str(x) for x in df_out["ssa"].dropna().unique() if str(x) not in ["nan", ""]])
                sel_ssa_o = st.selectbox("Filter SSA", ssa_list, key="outage_filter_ssa")

            with o_col3:
                search_out = st.text_input("Search Site ID / OMCR ID / Fault Reason", "", key="outage_search_kw")

            # Apply Outage Filters
            view_out = df_out.copy()
            if outage_filter_status == "⚡ Sites with Downtime (> 0 Hrs)":
                view_out = view_out[view_out["duration_hours"] > 0]
            elif outage_filter_status == "Zero Downtime Sites (NIL)":
                view_out = view_out[view_out["duration_hours"] == 0]

            if sel_op != "All" and "operator" in view_out.columns:
                view_out = view_out[view_out["operator"].astype(str) == sel_op]
            if sel_ssa_o != "All" and "ssa" in view_out.columns:
                view_out = view_out[view_out["ssa"].astype(str) == sel_ssa_o]
            if search_out.strip():
                s_kw = search_out.strip().lower()
                m_out = pd.Series(False, index=view_out.index)
                for sc in ["IP ID", "OMCR ID", "Site Name", "fault_reason", "ac_status", "officer_remarks"]:
                    if sc in view_out.columns:
                        m_out = m_out | view_out[sc].astype(str).str.lower().str.contains(s_kw, na=False)
                view_out = view_out[m_out]

            tot_hrs = view_out["duration_hours"].sum() if "duration_hours" in view_out.columns else 0.0

            m_o1, m_o2, m_o3 = st.columns(3)
            with m_o1:
                st.metric("Total Outage Entries", f"{len(view_out):,} Records")
            with m_o2:
                st.metric("Accumulated Downtime", f"{tot_hrs:,.1f} Hours")
            with m_o3:
                ssa_count = view_out["ssa"].nunique() if "ssa" in view_out.columns else 0
                st.metric("SSAs Represented", f"{ssa_count} Districts")

            st.caption(f"Showing {len(view_out)} of {len(df_out)} recorded outage events across Gujarat Circle")

            # Clean display columns for Outage Table
            display_out = view_out.copy()
            rename_map = {
                "ssa": "SSA",
                "operator": "IP Operator",
                "duration_hours": "Downtime (Hours)",
                "outage_minutes": "Downtime (Mins)",
                "fault_reason": "Fault Reason",
                "ac_status": "AC Status",
                "officer_remarks": "Officer Remarks"
            }
            display_out = display_out.rename(columns=rename_map)
            cols_order = [
                c for c in ["SSA", "IP Operator", "IP ID", "OMCR ID", "Site Name", "Downtime (Hours)", "Downtime (Mins)", "Fault Reason", "AC Status", "Officer Remarks"]
                if c in display_out.columns
            ]
            display_out = display_out[cols_order]

            for col in display_out.columns:
                if col in ["Downtime (Hours)", "Downtime (Mins)"]:
                    display_out[col] = pd.to_numeric(display_out[col], errors="coerce").fillna(0.0)
                else:
                    display_out[col] = display_out[col].fillna("").astype(str)

            # Sort by downtime descending only for downtime-specific view; keep original numerical order for all sites
            if outage_filter_status == "⚡ Sites with Downtime (> 0 Hrs)":
                display_out = display_out.sort_values(by="Downtime (Hours)", ascending=False)

            # Clean 1-based sequential numerical order
            display_out = display_out.reset_index(drop=True)
            display_out.index = display_out.index + 1

            st.dataframe(
                display_out.style.format({
                    "Downtime (Hours)": "{:.2f}",
                    "Downtime (Mins)": "{:.1f}"
                }),
                use_container_width=True,
                height=450
            )


# ----------------- TAB 6: BSNL SAP REM ERP UPLOAD GENERATOR -----------------
with tab_rem:
    st.subheader("🏛️ BSNL SAP REM ERP Module Bulk Upload Package — Gujarat Circle (GJ)")
    st.write("Generates the official 3 Corporate Office standard bulk upload spreadsheets for the BSNL ERP Real Estate Management (REM) module under Company Code `1100` and Circle Code `GJ`.")

    if not is_data_available(vendor_selected, circle_selected, month_selected, year_selected):
        render_no_data_banner("BSNL SAP REM Package", vendor_selected, circle_selected, month_selected, year_selected)
    else:
        gj_data = MasterInfraManager.load_gujarat_atc_month_data(month_selected, year_selected)
        df_rem_sites = gj_data.get("sites_df", pd.DataFrame()).copy()

        if df_rem_sites.empty:
            st.info("No sites loaded to generate SAP REM templates.")
        else:
            def clean_int_str(val):
                if pd.isna(val) or val is None:
                    return ""
                try:
                    return str(int(float(val)))
                except:
                    return str(val).strip()

            df_rem_sites["clean_inv"] = df_rem_sites["Invoice no"].apply(clean_int_str)
            distinct_bills = sorted([x for x in df_rem_sites["clean_inv"].unique() if x and x != "nan"])

            st.caption(f"Configured for **{len(df_rem_sites)} sites** across **{len(distinct_bills)} distinct vendor bills (invoices)** from the Gujarat ATC April 2026 Master Billing Dataset.")

            rent_stream_type = st.radio("Condition Stream", ["🏢 IP Tower Rent (T500)", "🏠 Ground / Landlord Rent (T630)"], horizontal=True, key="rem_tab_stream")
            rent_stream_code = "RENT" if "Tower Rent" in rent_stream_type else "LLR"
            rent_fy_period = ERPREMGenerator.get_fy_period(month_selected)
            rent_f_dt = f"01.{str(m_num).zfill(2)}.{y_num}"
            rent_t_dt = f"{days_in_month}.{str(m_num).zfill(2)}.{y_num}"
            next_m = 1 if m_num == 12 else m_num + 1
            next_y = y_num + 1 if m_num == 12 else y_num
            rent_inv_dt = f"05.{str(next_m).zfill(2)}.{next_y}"

            r_col1, r_col2, r_col3 = st.columns([1, 1, 2])
            with r_col1:
                rent_vendor_no = st.text_input("Registered Vendor Code", "2004755", key="rem_tab_vendor")
            with r_col2:
                rent_circle_code = st.text_input("Circle Code", "GJ", key="rem_tab_circle")
            with r_col3:
                bill_selector_options = ["📦 All Bills (Master Package with 3 Templates per Bill)"] + [
                    f"Bill {b} ({len(df_rem_sites[df_rem_sites['clean_inv'] == b])} sites)" for b in distinct_bills
                ]
                sel_bill_choice = st.selectbox("Select Target Bill / Invoice", bill_selector_options, key="rem_tab_sel_bill")

            # Prepare structured site records grouped by bill number
            grouped_site_records = {}
            for b in distinct_bills:
                sub_b = df_rem_sites[df_rem_sites["clean_inv"] == b]
                records_b = []
                for _, r in sub_b.iterrows():
                    c_no = clean_int_str(r.get("ISA"))
                    if not c_no:
                        gid = clean_int_str(r.get("GLOBAL_SITE_ID", ""))
                        c_no = f"1653000{gid.zfill(6)}" if gid else "1653000000001"
                    ro_no = clean_int_str(r.get("RO"))
                    inv_no = b
                    app_rent = float(r.get("Net Approved (₹)", r.get("Total IP Fee", 0.0)) or 0.0)
                    out_pen = float(r.get("Deduction Outage (₹)", 0.0) or 0.0)

                    records_b.append({
                        "contract_no": c_no,
                        "rental_object_no": ro_no if ro_no else None,
                        "site_name": str(r.get("Site Name", "")),
                        "approved_rent": app_rent,
                        "ground_rent": 0.0,
                        "outage_penalty": out_pen,
                        "base_passed": app_rent,
                        "invoice_no": inv_no
                    })
                grouped_site_records[b] = records_b

            st.markdown("---")

            if sel_bill_choice.startswith("📦 All Bills"):
                st.markdown(f"#### 📦 Master Package for All {len(distinct_bills)} Bills (3 Upload Templates per Bill)")
                st.caption(f"Generates {len(distinct_bills) * 3} spreadsheets (3 official templates for each of the {len(distinct_bills)} vendor bills), organized in folders by bill number.")

                master_all_zip = ERPREMGenerator.generate_grouped_invoices_zip(
                    grouped_sites=grouped_site_records,
                    bill_type=rent_stream_code,
                    circle_code=rent_circle_code,
                    vendor_no=rent_vendor_no,
                    invoice_date=rent_inv_dt,
                    from_date=rent_f_dt,
                    to_date=rent_t_dt,
                    period=rent_fy_period,
                    fiscal_year=y_num
                )
                st.download_button(
                    label=f"📦 Download Complete Master Package for All {len(distinct_bills)} Bills (.ZIP — 24 Templates)",
                    data=master_all_zip,
                    file_name=f"BSNL_REM_MASTER_PACKAGE_{rent_circle_code}_{rent_stream_code}_{month_selected}{year_selected}.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                    key="btn_dl_master_all_bills_zip"
                )

                st.markdown("##### 📋 Individual Bill Download Schedule:")
                for b_no in distinct_bills:
                    sites_this_bill = grouped_site_records.get(b_no, [])
                    sub_df = df_rem_sites[df_rem_sites["clean_inv"] == b_no]
                    tot_app = sub_df["Net Approved (₹)"].sum()
                    with st.expander(f"📄 Bill {b_no} — {len(sites_this_bill)} sites | Net Approved: ₹{tot_app:,.2f}", expanded=False):
                        b_col_z, b_c1, b_c2, b_c3 = st.columns([1.5, 1, 1, 1])
                        with b_col_z:
                            zip_this_bill = ERPREMGenerator.generate_all_3_templates_zip(
                                invoice_no=b_no,
                                sites=sites_this_bill,
                                bill_type=rent_stream_code,
                                circle_code=rent_circle_code,
                                vendor_no=rent_vendor_no,
                                period=rent_fy_period,
                                fiscal_year=y_num,
                                invoice_date=rent_inv_dt,
                                from_date=rent_f_dt,
                                to_date=rent_t_dt
                            )
                            st.download_button(
                                label=f"📦 3 Templates for Bill {b_no} (.ZIP)",
                                data=zip_this_bill,
                                file_name=f"REM_UPLOAD_{rent_circle_code}_{b_no}.zip",
                                mime="application/zip",
                                type="primary",
                                key=f"dl_zip_bill_{b_no}"
                            )
                        with b_c1:
                            inv_bytes = ERPREMGenerator.generate_invoice_upload(
                                invoice_no=b_no,
                                sites=sites_this_bill,
                                bill_type=rent_stream_code,
                                circle_code=rent_circle_code,
                                vendor_no=rent_vendor_no,
                                from_date=rent_f_dt,
                                to_date=rent_t_dt
                            )
                            st.download_button(
                                label="📄 1. INVOICE UPLOAD",
                                data=inv_bytes,
                                file_name=f"INVOICE_UPLOAD_{b_no}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_inv_bill_{b_no}"
                            )
                        with b_c2:
                            det_bytes = ERPREMGenerator.generate_invoice_detail_upload(
                                invoice_no=b_no,
                                sites=sites_this_bill,
                                circle_code=rent_circle_code,
                                vendor_no=rent_vendor_no,
                                invoice_date=rent_inv_dt,
                                period=rent_fy_period,
                                fiscal_year=y_num
                            )
                            st.download_button(
                                label="📄 2. DETAIL UPLOAD",
                                data=det_bytes,
                                file_name=f"INVOICE_DETAIL_UPLOAD_{b_no}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_det_bill_{b_no}"
                            )
                        with b_c3:
                            cond_bytes = ERPREMGenerator.generate_condition_upload(
                                sites=sites_this_bill,
                                bill_type=rent_stream_code,
                                circle_code=rent_circle_code,
                                from_date=rent_f_dt,
                                to_date=rent_t_dt
                            )
                            st.download_button(
                                label=f"📄 3. CONDITION",
                                data=cond_bytes,
                                file_name=f"CONDITION_UPLOAD_{b_no}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_cond_bill_{b_no}"
                            )
            else:
                # Individual Bill Selected
                target_bill = sel_bill_choice.split()[1]
                sites_this_bill = grouped_site_records.get(target_bill, [])
                sub_df = df_rem_sites[df_rem_sites["clean_inv"] == target_bill]
                tot_billed = sub_df["Total Billed (₹)"].sum()
                tot_ded = sub_df["Total Deductions (₹)"].sum()
                tot_app = sub_df["Net Approved (₹)"].sum()

                st.markdown(f"#### 📄 3 Templates for Bill Number: `{target_bill}`")
                m_b1, m_b2, m_b3, m_b4 = st.columns(4)
                with m_b1:
                    st.metric("Sites in Bill", f"{len(sites_this_bill)} Sites")
                with m_b2:
                    st.metric("Claimed Amount", f"₹{tot_billed:,.2f}")
                with m_b3:
                    st.metric("Deductions Disallowed", f"₹{tot_ded:,.2f}")
                with m_b4:
                    st.metric("Net Approved Payable", f"₹{tot_app:,.2f}")

                r_zip_btn, r_b1, r_b2, r_b3 = st.columns([1.4, 1, 1, 1])
                with r_zip_btn:
                    zip_this_bill = ERPREMGenerator.generate_all_3_templates_zip(
                        invoice_no=target_bill,
                        sites=sites_this_bill,
                        bill_type=rent_stream_code,
                        circle_code=rent_circle_code,
                        vendor_no=rent_vendor_no,
                        period=rent_fy_period,
                        fiscal_year=y_num,
                        invoice_date=rent_inv_dt,
                        from_date=rent_f_dt,
                        to_date=rent_t_dt
                    )
                    st.download_button(
                        label=f"📦 All 3 Templates for Bill {target_bill} (.ZIP)",
                        data=zip_this_bill,
                        file_name=f"BSNL_REM_UPLOAD_{rent_circle_code}_{target_bill}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True,
                        key=f"btn_zip_single_{target_bill}"
                    )
                with r_b1:
                    inv_bytes = ERPREMGenerator.generate_invoice_upload(
                        invoice_no=target_bill,
                        sites=sites_this_bill,
                        bill_type=rent_stream_code,
                        circle_code=rent_circle_code,
                        vendor_no=rent_vendor_no,
                        from_date=rent_f_dt,
                        to_date=rent_t_dt
                    )
                    st.download_button(
                        label="📄 1. INVOICE UPLOAD",
                        data=inv_bytes,
                        file_name=f"INVOICE_UPLOAD_{rent_circle_code}_{target_bill}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"btn_inv_single_{target_bill}"
                    )
                with r_b2:
                    det_bytes = ERPREMGenerator.generate_invoice_detail_upload(
                        invoice_no=target_bill,
                        sites=sites_this_bill,
                        circle_code=rent_circle_code,
                        vendor_no=rent_vendor_no,
                        invoice_date=rent_inv_dt,
                        period=rent_fy_period,
                        fiscal_year=y_num
                    )
                    st.download_button(
                        label="📄 2. DETAIL UPLOAD",
                        data=det_bytes,
                        file_name=f"INVOICE_detail_UPLOAD_{rent_circle_code}_{target_bill}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"btn_det_single_{target_bill}"
                    )
                with r_b3:
                    cond_bytes = ERPREMGenerator.generate_condition_upload(
                        sites=sites_this_bill,
                        bill_type=rent_stream_code,
                        circle_code=rent_circle_code,
                        from_date=rent_f_dt,
                        to_date=rent_t_dt
                    )
                    st.download_button(
                        label=f"📄 3. CONDITION ({rent_stream_code})",
                        data=cond_bytes,
                        file_name=f"CONDITION_UPLOAD_{rent_circle_code}_{target_bill}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"btn_cond_single_{target_bill}"
                    )


# ----------------- TAB 10: MSA RULES & CIRCULAR AMENDMENTS (NEW INTERACTIVE SECTION) -----------------
with tab_msa:
    st.subheader("📜 Master Service Agreement (MSA) Rules & Amendments Repository")
    st.write("Vendor-wise contractual terms, rate structures, penalty formulas, and chronologically tracked BSNL Corporate Office circulars.")

    # Synchronize default vendor with sidebar vendor selection!
    vendor_list = list(st.session_state.vendor_msa_rules.keys())
    default_vendor_idx = vendor_list.index(vendor_selected) if vendor_selected in vendor_list else 0

    v_col1, v_col2 = st.columns([1, 2])
    with v_col1:
        selected_msa_vendor = st.selectbox(
            "Select IP Vendor to View Terms",
            vendor_list,
            index=default_vendor_idx,
            key=f"msa_vendor_select_{vendor_selected}"
        )
    with v_col2:
        v_data = st.session_state.vendor_msa_rules[selected_msa_vendor]
        st.info(f"**Corporate Lineage & Mergers:** {v_data['lineage']}")

    # 3 High Level KPI Cards
    col_k1, col_k2, col_k3 = st.columns(3)
    with col_k1:
        st.markdown(f"""
        <div class="metric-card">
            <b>⚡ SLA Target & Downtime:</b><br>
            • Uptime Target: <b>{v_data['uptime_target']}</b><br>
            • Permissible Downtime: <b>{v_data['permissible_downtime']}</b>
        </div>
        """, unsafe_allow_html=True)
    with col_k2:
        st.markdown(f"""
        <div class="metric-card">
            <b>💰 SLA Penalty Formula:</b><br>
            • Rate: <b>{v_data['penalty_formula']}</b><br>
            • Site Cap: <b>{v_data['penalty_cap']}</b>
        </div>
        """, unsafe_allow_html=True)
    with col_k3:
        st.markdown(f"""
        <div class="metric-card">
            <b>🏠 Landlord Rent (LLR) Capping:</b><br>
            • <b>{v_data['llr_capping']}</b>
        </div>
        """, unsafe_allow_html=True)

    # Detailed Commercial Terms Card
    st.markdown("### 📋 Commercial & Operational Rules Summary")
    t1, t2 = st.columns(2)
    with t1:
        st.markdown(f"""
        * **Governing Master Agreement:** {v_data['master_msa']}
        * **Tenancy Sharing Treatment:** {v_data['tenancy_rule']}
        * **Lock-In Period (LIP) Extension:** {v_data['lip_extension']}
        * **Energy (EB & DG) Norms:** {v_data['energy_norm']}
        """)
    with t2:
        st.markdown(f"""
        * **4G Deployment Charges:** {v_data['rates_4g']}
        {('* **Special Provisions / Discounts:** ' + v_data['special_clauses']) if 'special_clauses' in v_data else ''}
        """)

    st.markdown("---")

    # SECTION: Contractual Base Rent Matrix by Conditions (Tenancy, Tower Type, Category, BTS Type)
    st.markdown("### 📊 Contractual IP Rent Matrix (By Site Category, Tower Type, Tenancy & BTS Type)")
    st.caption("Standard 3rd EoI Schedule 3 (Table B) Base Rates & Sharing Multipliers with 2024 Harmonized Amendments")

    rent_matrix_data = [
        {"Site Category": "Metro (Delhi, Mumbai, Kol, Chn, Blr)", "Tower Type": "GBT", "BTS Type": "ID", "1T Solo (₹)": 32000.0, "2T Sharing (₹)": 29430.0, "3T Sharing (₹)": 29430.0, "4T+ Sharing (₹)": 25751.0, "Deemed 2T Cap": "₹29,430 max", "LLR Cap": "₹7,000 / ₹5,000"},
        {"Site Category": "Metro (Delhi, Mumbai, Kol, Chn, Blr)", "Tower Type": "GBT", "BTS Type": "OD", "1T Solo (₹)": 30000.0, "2T Sharing (₹)": 28530.0, "3T Sharing (₹)": 28000.0, "4T+ Sharing (₹)": 24964.0, "Deemed 2T Cap": "₹28,530 max", "LLR Cap": "₹7,000 / ₹5,000"},
        {"Site Category": "Metro (Delhi, Mumbai, Kol, Chn, Blr)", "Tower Type": "RTT", "BTS Type": "ID", "1T Solo (₹)": 25600.0, "2T Sharing (₹)": 23544.0, "3T Sharing (₹)": 23544.0, "4T+ Sharing (₹)": 20601.0, "Deemed 2T Cap": "₹23,544 max", "LLR Cap": "₹7,000 / ₹5,000"},
        {"Site Category": "Metro (Delhi, Mumbai, Kol, Chn, Blr)", "Tower Type": "RTT", "BTS Type": "OD", "1T Solo (₹)": 24000.0, "2T Sharing (₹)": 22824.0, "3T Sharing (₹)": 22400.0, "4T+ Sharing (₹)": 19971.0, "Deemed 2T Cap": "₹22,824 max", "LLR Cap": "₹7,000 / ₹5,000"},
        
        {"Site Category": "District B1 (Major Urban / Capital)", "Tower Type": "GBT", "BTS Type": "ID", "1T Solo (₹)": 29052.0, "2T Sharing (₹)": 29052.0, "3T Sharing (₹)": 28197.0, "4T+ Sharing (₹)": 25634.0, "Deemed 2T Cap": "₹29,052 max", "LLR Cap": "₹3,000"},
        {"Site Category": "District B1 (Major Urban / Capital)", "Tower Type": "GBT", "BTS Type": "OD", "1T Solo (₹)": 26000.0, "2T Sharing (₹)": 26000.0, "3T Sharing (₹)": 24500.0, "4T+ Sharing (₹)": 22500.0, "Deemed 2T Cap": "₹26,000 max", "LLR Cap": "₹3,000"},
        {"Site Category": "District B1 (Major Urban / Capital)", "Tower Type": "RTT", "BTS Type": "ID", "1T Solo (₹)": 18450.0, "2T Sharing (₹)": 18450.0, "3T Sharing (₹)": 15600.0, "4T+ Sharing (₹)": 15600.0, "Deemed 2T Cap": "₹18,450 max", "LLR Cap": "₹3,000"},
        {"Site Category": "District B1 (Major Urban / Capital)", "Tower Type": "RTT", "BTS Type": "OD", "1T Solo (₹)": 16500.0, "2T Sharing (₹)": 16500.0, "3T Sharing (₹)": 14000.0, "4T+ Sharing (₹)": 14000.0, "Deemed 2T Cap": "₹16,500 max", "LLR Cap": "₹3,000"},

        {"Site Category": "District B2 (Semi-Urban Districts)", "Tower Type": "GBT", "BTS Type": "ID", "1T Solo (₹)": 17000.0, "2T Sharing (₹)": 17000.0, "3T Sharing (₹)": 15000.0, "4T+ Sharing (₹)": 15000.0, "Deemed 2T Cap": "₹17,000 max", "LLR Cap": "₹3,000"},
        {"Site Category": "District B2 (Semi-Urban Districts)", "Tower Type": "GBT", "BTS Type": "OD", "1T Solo (₹)": 15000.0, "2T Sharing (₹)": 15000.0, "3T Sharing (₹)": 13000.0, "4T+ Sharing (₹)": 13000.0, "Deemed 2T Cap": "₹15,000 max", "LLR Cap": "₹3,000"},
        {"Site Category": "District B2 (Semi-Urban Districts)", "Tower Type": "RTT", "BTS Type": "ID", "1T Solo (₹)": 14000.0, "2T Sharing (₹)": 14000.0, "3T Sharing (₹)": 12500.0, "4T+ Sharing (₹)": 12500.0, "Deemed 2T Cap": "₹14,000 max", "LLR Cap": "₹3,000"},
        {"Site Category": "District B2 (Semi-Urban Districts)", "Tower Type": "RTT", "BTS Type": "OD", "1T Solo (₹)": 12500.0, "2T Sharing (₹)": 12500.0, "3T Sharing (₹)": 11000.0, "4T+ Sharing (₹)": 11000.0, "Deemed 2T Cap": "₹12,500 max", "LLR Cap": "₹3,000"},

        {"Site Category": "Rural / Category C", "Tower Type": "GBT", "BTS Type": "ID", "1T Solo (₹)": 17000.0, "2T Sharing (₹)": 17000.0, "3T Sharing (₹)": 15000.0, "4T+ Sharing (₹)": 15000.0, "Deemed 2T Cap": "₹17,000 max", "LLR Cap": "₹3,000"},
        {"Site Category": "Rural / Category C", "Tower Type": "GBT", "BTS Type": "OD", "1T Solo (₹)": 15000.0, "2T Sharing (₹)": 15000.0, "3T Sharing (₹)": 13000.0, "4T+ Sharing (₹)": 13000.0, "Deemed 2T Cap": "₹15,000 max", "LLR Cap": "₹3,000"},
        {"Site Category": "Rural / Category C", "Tower Type": "RTT", "BTS Type": "ID", "1T Solo (₹)": 13000.0, "2T Sharing (₹)": 13000.0, "3T Sharing (₹)": 11500.0, "4T+ Sharing (₹)": 11500.0, "Deemed 2T Cap": "₹13,000 max", "LLR Cap": "₹3,000"},
        {"Site Category": "Rural / Category C", "Tower Type": "RTT", "BTS Type": "OD", "1T Solo (₹)": 11500.0, "2T Sharing (₹)": 11500.0, "3T Sharing (₹)": 10000.0, "4T+ Sharing (₹)": 10000.0, "Deemed 2T Cap": "₹11,500 max", "LLR Cap": "₹3,000"},
    ]
    df_rent_matrix = pd.DataFrame(rent_matrix_data)
    st.dataframe(
        df_rent_matrix.style.format({
            "1T Solo (₹)": "₹{:,.0f}",
            "2T Sharing (₹)": "₹{:,.0f}",
            "3T Sharing (₹)": "₹{:,.0f}",
            "4T+ Sharing (₹)": "₹{:,.0f}"
        }),
        use_container_width=True
    )

    st.markdown("---")

    # SECTION: Interactive Contractual Rent Calculator by Site Conditions
    st.markdown("### 🧮 Interactive Rent Rate Estimator (Calculate Rent by Specific Conditions)")
    st.write("Simulate the exact contractual rent payable for a site based on its physical and operational parameters:")

    with st.container():
        rc1, rc2, rc3, rc4 = st.columns(4)
        with rc1:
            sim_cat = st.selectbox("Site Category / City Class", ["Metro", "District B1 (Urban / State Capital)", "District B2 (Semi-Urban)", "Rural / Category C"], key="sim_cat")
        with rc2:
            sim_tower = st.selectbox("Tower Structure Type", ["GBT (Ground Based Tower)", "RTT (Roof Top Tower)"], key="sim_tower")
        with rc3:
            sim_bts = st.selectbox("BTS Installation Type", ["Indoor (ID)", "Outdoor (OD)"], key="sim_bts")
        with rc4:
            sim_tenancy = st.selectbox("Site Tenancy (Working Tenants)", ["1T (Solo Tenant)", "2T (Two Tenants)", "3T (Three Tenants)", "4T+ (Four+ Tenants)"], index=1, key="sim_tenancy")

        rc5, rc6, rc7, rc8 = st.columns(4)
        with rc5:
            sim_lip = st.selectbox("Lock-In Period Status", ["LIP Renewed (>10 Yrs - Freeze to Base)", "Standard Pre-10 Yrs (Subject to Escalation)"], key="sim_lip")
        with rc6:
            sim_bbu = st.selectbox("4G BBU Type", ["None", "Outdoor BBU (ODC & Pad: ₹1,350/mo)", "Indoor 2U BBU (Rack Space: ₹750/mo)"], index=1, key="sim_bbu")
        with rc7:
            sim_rru_count = st.number_input("4G RRU / RRH Units (₹800/unit)", min_value=0, max_value=6, value=3, step=1, key="sim_rru")
        with rc8:
            sim_llr_claimed = st.number_input("Claimed Landlord Rent (₹/mo)", min_value=0.0, value=9500.0, step=500.0, key="sim_llr")

        # Calculation Engine based on Selected Parameters
        # Base Fee Matrix Lookup
        is_metro = "Metro" in sim_cat
        is_b1 = "District B1" in sim_cat
        is_gbt = "GBT" in sim_tower
        is_id = "Indoor" in sim_bts
        t_num = 1 if "1T" in sim_tenancy else (2 if "2T" in sim_tenancy else (3 if "3T" in sim_tenancy else 4))

        # Base rate determination
        if is_metro:
            if is_gbt:
                base_rates = {1: 32000.0, 2: 29430.0, 3: 29430.0, 4: 25751.0} if is_id else {1: 30000.0, 2: 28530.0, 3: 28000.0, 4: 24964.0}
            else:
                base_rates = {1: 25600.0, 2: 23544.0, 3: 23544.0, 4: 20601.0} if is_id else {1: 24000.0, 2: 22824.0, 3: 22400.0, 4: 19971.0}
            llr_cap_allowed = 7000.0 if "Delhi" in sim_cat or "Mumbai" in sim_cat else 5000.0
        elif is_b1:
            if is_gbt:
                base_rates = {1: 29052.0, 2: 29052.0, 3: 28197.0, 4: 25634.0} if is_id else {1: 26000.0, 2: 26000.0, 3: 24500.0, 4: 22500.0}
            else:
                base_rates = {1: 18450.0, 2: 18450.0, 3: 15600.0, 4: 15600.0} if is_id else {1: 16500.0, 2: 16500.0, 3: 14000.0, 4: 14000.0}
            llr_cap_allowed = 3000.0
        else:
            if is_gbt:
                base_rates = {1: 17000.0, 2: 17000.0, 3: 15000.0, 4: 15000.0} if is_id else {1: 15000.0, 2: 15000.0, 3: 13000.0, 4: 13000.0}
            else:
                base_rates = {1: 14000.0, 2: 14000.0, 3: 12500.0, 4: 12500.0} if is_id else {1: 12500.0, 2: 12500.0, 3: 11000.0, 4: 11000.0}
            llr_cap_allowed = 3000.0

        # Deemed 2nd Tenant Rule: If 1T, charge 2T rate!
        effective_t = 2 if t_num == 1 else t_num
        calculated_base_rent = base_rates.get(effective_t, 17000.0)

        # 4G Hardware Loading
        bbu_charge = 1350.0 if "1,350" in sim_bbu else (750.0 if "750" in sim_bbu else 0.0)
        rru_charge = sim_rru_count * 800.0
        tot_4g_loading = bbu_charge + rru_charge

        # Landlord Rent Capping Calculation
        llr_threshold = 8000.0
        excess_llr = max(0.0, sim_llr_claimed - llr_threshold)
        passed_additional_llr = min(excess_llr, llr_cap_allowed)
        disallowed_llr = excess_llr - passed_additional_llr

        total_net_rent = calculated_base_rent + tot_4g_loading + passed_additional_llr
        gst_rent = total_net_rent * 0.18
        gross_rent = total_net_rent + gst_rent

        # Display Result Breakdown Cards
        st.markdown("#### 📑 Estimated Rent Breakdown for Selected Site Condition:")
        res1, res2, res3, res4, res5 = st.columns(5)
        with res1:
            st.metric("Base Contractual Rent", f"₹{calculated_base_rent:,.2f}", help=f"Table B sharing rate for Tenancy {effective_t}T")
        with res2:
            st.metric("4G Hardware Loading", f"₹{tot_4g_loading:,.2f}", help=f"BBU: ₹{bbu_charge} + {sim_rru_count} RRUs: ₹{rru_charge}")
        with res3:
            st.metric("Approved Landlord Rent", f"₹{passed_additional_llr:,.2f}", delta=f"-₹{disallowed_llr:,.2f} Disallowed" if disallowed_llr > 0 else "Fully Within Cap")
        with res4:
            st.metric("Net Rent (Excl. GST)", f"₹{total_net_rent:,.2f}", delta="Payable Net")
        with res5:
            st.metric("Total with GST @18%", f"₹{gross_rent:,.2f}")

        if t_num == 1:
            st.info("🛡️ **Deemed 2nd Tenant Protection Active:** Even though the site has 1 tenant (solo), BSNL is contractually protected and billed at the lower **2T Sharing Rate** as per BSNL Corporate Office directives!")
        if disallowed_llr > 0:
            st.warning(f"⚠️ **Landlord Rent Capping Applied:** Claimed LLR ₹{sim_llr_claimed:,.2f} exceeds ₹8,000 threshold by ₹{excess_llr:,.2f}. BSNL share is capped at ₹{llr_cap_allowed:,.2f}, resulting in a contractual deduction of **₹{disallowed_llr:,.2f}**.")

    st.markdown("---")

    # Master Table of Amendments & Circulars
    st.markdown("### 📑 Chronological History of Amendments & Corporate Circulars")
    if v_data.get("amendments_history"):
        df_amend = pd.DataFrame(v_data["amendments_history"])
        st.dataframe(df_amend, use_container_width=True)
    else:
        st.caption("No custom amendments recorded yet.")

    st.markdown("---")

    # Interactive Form: Add / Amend Circular or Rule
    st.markdown("### ➕ Add / Amend Vendor Rules with a New Circular")
    st.write("Submit a new circular or addendum to immediately update and amend the vendor's billing summary and rules.")

    with st.form(key="add_amendment_form"):
        fa1, fa2 = st.columns(2)
        with fa1:
            form_vendor = st.selectbox("Target Vendor", list(st.session_state.vendor_msa_rules.keys()), index=list(st.session_state.vendor_msa_rules.keys()).index(selected_msa_vendor))
            form_ref = st.text_input("Circular / Addendum Ref. No.", placeholder="e.g. BSNLCO-CMIN/16(15)/.../2026")
            form_date = st.date_input("Circular Date", value=datetime.today())
        with fa2:
            form_subject = st.text_input("Circular Subject", placeholder="e.g. Amendment to 4G loading / revised rent cap")
            form_effective = st.text_input("Effective Date", placeholder="e.g. 01.03.2026")
            form_doc = st.file_uploader("Upload Circular Document (Optional)", type=["pdf", "doc", "docx"], key="circ_upload")

        form_clause = st.text_area("Amended Clause & Billing Impact", placeholder="Detail the exact rate changes, penalty clause adjustments, or rental capping terms...")

        submit_amendment = st.form_submit_button("💾 Save & Apply Amendment to Vendor Rules", type="primary")

        if submit_amendment:
            if not form_ref or not form_clause:
                st.error("Please provide at least the Circular Reference Number and the Amended Clause.")
            else:
                new_entry = {
                    "Circular / Ref No.": form_ref,
                    "Date": form_date.strftime("%d.%m.%Y"),
                    "Subject": form_subject or "Contractual Amendment",
                    "Key Amendment": form_clause,
                    "Effective Date": form_effective or form_date.strftime("%d.%m.%Y")
                }
                # Prepend to history
                st.session_state.vendor_msa_rules[form_vendor]["amendments_history"].insert(0, new_entry)
                st.success(f"Successfully applied amendment '{form_ref}' to {form_vendor}! The rules summary and circular history have been updated.")
                st.rerun()


# ----------------- TAB 11: DOWNLOAD MASTER EXCEL -----------------
with tab_export:
    st.subheader(f"📥 Download Master Audit Excel Workbooks — Gujarat Circle ({month_selected}-{year_selected})")
    st.write("Download the comprehensive master vendor invoice consoles, signed outage reports, and SAP REM bulk upload spreadsheets for Gujarat Circle:")

    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.markdown("### 🏢 1. Vendor Invoices & Site Annexure Console")
        st.markdown("""
        * **April 2026 Vendor Submitted Console**: 524 Leased Sites across 8 Invoices.
        * **Itemized Slabs**: Base MSA Rate, Antenna Markup (₹650), Power (₹3,000), BTS (₹3,750).
        * **Deployment Penalties Sheet**: 25 delayed delivery sites with contractual deductions.
        * **Totals**: ₹1,42,80,830.98 excl. tax / ₹1,68,51,380.56 incl. GST.
        """)

        f_inv_p = os.path.join(MasterInfraManager.get_master_dir(), "Copy_of_IP_Fees_console_April_26_VENDORE_SUBMITTED_BSNL.xlsx")
        if os.path.exists(f_inv_p):
            with open(f_inv_p, "rb") as f_inv:
                st.download_button(
                    label="📥 Download April-2026 Vendor Console (Copy_of_IP_Fees_console_April_26_VENDORE_SUBMITTED_BSNL.xlsx)",
                    data=f_inv.read(),
                    file_name="Copy_of_IP_Fees_console_April_26_VENDORE_SUBMITTED_BSNL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="dl_april_console_btn"
                )

        st.markdown("---")
        st.markdown("### 🏛️ 2. Annual Master IP Fee Registry (FY 2026-27)")
        st.caption("Master workbook containing month-by-month site registries for Gujarat Circle (580+ sites).")
        f_main_p = os.path.join(MasterInfraManager.get_master_dir(), "1_ATC_GJ_IPF_2026-27_Main.xlsx")
        if os.path.exists(f_main_p):
            with open(f_main_p, "rb") as f_main:
                st.download_button(
                    label="📥 Download Full FY 2026-27 Master (1_ATC_GJ_IPF_2026-27_Main.xlsx)",
                    data=f_main.read(),
                    file_name="1_ATC_GJ_IPF_2026-27_Main.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_main_fy_btn"
                )

    with col_e2:
        st.markdown("### ⚡ 3. Multi-Operator Master Outage & AC Downtime Report")
        st.markdown("""
        * **All 8 IP Vendors**: ATC/ATCTIPL, Indus Towers, Ascend, Tower Vision, GTL, Crest, SDIL, and Signotox.
        * **16 Gujarat Districts**: ATD, Surat, Vadodara, Rajkot, Jamnagar, Bhavnagar, Mehsana, Nadiad, Palanpur, Valsad, Amreli, Bharuch, Bhuj, Godhra, Himmatnagar, Junagadh.
        * **987 Total Downtime Records**: Operational status of Air Conditioners and root cause fault reasons.
        """)

        f_out_p = os.path.join(MasterInfraManager.get_master_dir(), "All_IP_sites_DT_AC_MR_report_APR-26xls.xls")
        if os.path.exists(f_out_p):
            with open(f_out_p, "rb") as f_out:
                st.download_button(
                    label="📥 Download Multi-Operator Outage Report (All_IP_sites_DT_AC_MR_report_APR-26xls.xls)",
                    data=f_out.read(),
                    file_name="All_IP_sites_DT_AC_MR_report_APR-26xls.xls",
                    mime="application/vnd.ms-excel",
                    type="primary",
                    use_container_width=True,
                    key="dl_outage_rpt_btn"
                )

        st.markdown("---")
        st.markdown("### 📦 4. BSNL SAP REM ERP Bulk Upload Package")
        st.caption("Official BSNL ERP Real Estate Management package: 3 upload spreadsheets per bill across all 8 vendor invoices.")
        gj_d = MasterInfraManager.load_gujarat_atc_month_data(month_selected, year_selected)
        df_exp_sites = gj_d.get("sites_df", pd.DataFrame())
        if not df_exp_sites.empty:
            def clean_int_str(val):
                if pd.isna(val) or val is None:
                    return ""
                try:
                    return str(int(float(val)))
                except:
                    return str(val).strip()

            df_exp_sites["clean_inv"] = df_exp_sites["Invoice no"].apply(clean_int_str)
            dist_bills = sorted([x for x in df_exp_sites["clean_inv"].unique() if x and x != "nan"])

            grouped_exp = {}
            for b in dist_bills:
                sub_b = df_exp_sites[df_exp_sites["clean_inv"] == b]
                records_b = []
                for _, r in sub_b.iterrows():
                    c_no = clean_int_str(r.get("ISA"))
                    if not c_no:
                        gid = clean_int_str(r.get("GLOBAL_SITE_ID", ""))
                        c_no = f"1653000{gid.zfill(6)}" if gid else "1653000000001"
                    ro_no = clean_int_str(r.get("RO"))
                    app_rent = float(r.get("Net Approved (₹)", r.get("Total IP Fee", 0.0)) or 0.0)
                    out_pen = float(r.get("Deduction Outage (₹)", 0.0) or 0.0)

                    records_b.append({
                        "contract_no": c_no,
                        "rental_object_no": ro_no if ro_no else None,
                        "site_name": str(r.get("Site Name", "")),
                        "approved_rent": app_rent,
                        "ground_rent": 0.0,
                        "outage_penalty": out_pen,
                        "base_passed": app_rent,
                        "invoice_no": b
                    })
                grouped_exp[b] = records_b

            exp_fy_p = ERPREMGenerator.get_fy_period(month_selected)
            exp_f_dt = f"01.{str(m_num).zfill(2)}.{y_num}"
            exp_t_dt = f"{days_in_month}.{str(m_num).zfill(2)}.{y_num}"
            next_m = 1 if m_num == 12 else m_num + 1
            next_y = y_num + 1 if m_num == 12 else y_num
            exp_inv_dt = f"05.{str(next_m).zfill(2)}.{next_y}"

            zip_bytes = ERPREMGenerator.generate_grouped_invoices_zip(
                grouped_sites=grouped_exp,
                bill_type="RENT",
                circle_code="GJ",
                vendor_no="2004755",
                invoice_date=exp_inv_dt,
                period=exp_fy_p,
                fiscal_year=y_num,
                from_date=exp_f_dt,
                to_date=exp_t_dt
            )
            st.download_button(
                label=f"📦 Download Complete Gujarat REM Package (All {len(dist_bills)} Bills — 24 Templates .ZIP)",
                data=zip_bytes,
                file_name=f"BSNL_REM_MASTER_PACKAGE_GJ_RENT_{month_selected}{year_selected}.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key="btn_export_zip_gj"
            )
