# BSNL Telecom Infrastructure (IP-1) Billing & SLA Reconciliation System

A dedicated software solution designed for **BSNL (Bharat Sanchar Nigam Limited)** to automate monthly billing verification, site parameter audit, and SLA outage penalty reconciliation against Master Service Agreements (MSAs) with Infrastructure Providers (e.g., **Elevar Digitel / ATC, Indus Towers, Ascend Telecom, Suyog, Summit Digitel**).

---

## 🎯 What This Software Does

1. **Automated Site Annexure Ingestion:**
   * Converts monthly vendor billing files (PDF / Excel) into standardized site matrices.
   * Audits site parameters: **Global ID, Vendor Site ID, Tenancy (Anchor vs Sharing), City Class (Metro, Dist, Rural), Tower Type (GBT, RTT), and BTS Type (Indoor/Outdoor)**.

2. **Master Service Agreement (MSA) SLA & Outage Penalty Engine:**
   * Reads outage minutes per site from field sign-off reports (PDF / Excel) with **RapidOCR** support for scanned/stamped sheets.
   * Calculates permissible monthly downtime threshold based on days in the month (e.g., 28 days = **40.32 mins**, 30 days = **43.20 mins** at 99.90% SLA).
   * Applies the exact contractual penalty formula (**₹500/hour or part thereof**, capped at **₹10,000 maximum per site per month** as per MSA Clause 7.0).

3. **Energy & Power Pass-Through Recalculation:**
   * Re-evaluates Diesel Generator (DG) and Electricity Board (EB) charges according to HPC approved norms ($DG = \text{Run Hrs} \times CPH \times \text{Fuel Rate}$, $EB = (24 - DG) \times \text{Rate} \times CPH$).
   * Enforces BSNL standard billing policy: $\text{Passed Base Amount} = \min(\text{Claimed}, \text{Calculated})$.

4. **Multi-Tab Excel Audit Deliverables:**
   * **Tab 1: Executive Summary & MIRO:** Total claimed vs BSNL approved vs net deductions, ready for SAP ERP MIRO posting.
   * **Tab 2: Site-Wise Reconciliation:** All 500+ sites color-coded (Green = Verified, Red = Disputed/Deduction, Yellow = Warning).
   * **Tab 3: Outage & SLA Audit:** Min-by-min downtime calculations and field officer notes (e.g., cable theft, power cuts).
   * **Tab 4: Vendor Debit Note Annexure:** Itemized list of overbilled sites ready to submit to tower vendors for credit note issuance.

---

## 🚀 How to Launch and Use

### Quick Launch on Windows
Double-click:
```bat
run_app.bat
```
This automatically starts the Streamlit dashboard in your web browser.

### Or via Terminal / Command Prompt:
```powershell
python -m streamlit run app.py
```

---

## 📁 System Architecture

```
papa/
├── app.py                      # Interactive Web Dashboard (Streamlit)
├── run_app.bat                 # 1-Click Windows desktop launcher
├── engine/
│   ├── msa_rules.py            # MSA rate cards, SLA formulas, penalty cap
│   ├── invoice_parser.py       # PDF invoice & Excel site annexure extractor
│   ├── outage_parser.py        # Outage sign-off parser (with RapidOCR)
│   ├── reconciler.py           # Multi-factor audit & variance engine
│   └── excel_generator.py      # Professional 4-tab styled Excel report builder
├── BSNL_Audit_Report_FEB-26.xlsx  # Sample generated output report
└── README.md
```
