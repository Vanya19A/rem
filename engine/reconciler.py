"""
Core Reconciliation Engine for BSNL Telecom Infrastructure Invoices
Performs multi-factor audit:
1. Tenancy validation (Claimed vs Field Sign-off)
2. SLA Outage Penalty Computation (as per MSA clauses)
3. Energy HPC norm recalculation (EB + DG vs Vendor Claimed)
4. Discrepancy & Dispute Matrix Generation
"""

from typing import Dict, List, Any, Optional
import pandas as pd
import numpy as np

from .msa_rules import (
    calculate_outage_penalty,
    calculate_energy_charges,
    calculate_allowed_downtime_minutes,
    DEFAULT_UPTIME_TARGET
)


class ReconciliationEngine:
    def __init__(self, days_in_month: int = 28, uptime_target: float = DEFAULT_UPTIME_TARGET):
        self.days_in_month = days_in_month
        self.uptime_target = uptime_target
        self.allowed_downtime = calculate_allowed_downtime_minutes(days_in_month, uptime_target)

    def reconcile(
        self,
        site_df: pd.DataFrame,
        outage_df: Optional[pd.DataFrame] = None,
        master_df: Optional[pd.DataFrame] = None
    ) -> pd.DataFrame:
        """
        Merge and reconcile site annexure with outage and master data.
        """
        df = site_df.copy()

        # Ensure global_id is string
        df["global_id"] = df["global_id"].astype(str)

        # Merge Outage Data if available
        if outage_df is not None and not outage_df.empty:
            outage_df = outage_df.copy()
            outage_df["global_id"] = outage_df["global_id"].astype(str)
            
            # Select relevant columns from outage_df
            merge_cols = ["global_id", "outage_minutes"]
            if "verified_tenancy" in outage_df.columns:
                merge_cols.append("verified_tenancy")
            if "field_remark" in outage_df.columns:
                merge_cols.append("field_remark")

            sub_outage = outage_df[merge_cols].drop_duplicates(subset=["global_id"])
            
            # If outage_minutes already in df, preserve or update
            if "outage_minutes" in df.columns:
                df = df.rename(columns={"outage_minutes": "inv_outage_minutes"})
                df = df.merge(sub_outage, on="global_id", how="left")
                df["outage_minutes"] = df["outage_minutes"].fillna(df["inv_outage_minutes"]).fillna(0.0)
            else:
                df = df.merge(sub_outage, on="global_id", how="left")
                df["outage_minutes"] = df["outage_minutes"].fillna(0.0)
        else:
            if "outage_minutes" not in df.columns:
                df["outage_minutes"] = 0.0
            else:
                df["outage_minutes"] = pd.to_numeric(df["outage_minutes"], errors="coerce").fillna(0.0)

        # Reconcile Tenancy
        if "working_tenancy" not in df.columns:
            if "verified_tenancy" in df.columns:
                df["working_tenancy"] = df["verified_tenancy"].fillna(df.get("claimed_tenancy", 2.0))
            else:
                df["working_tenancy"] = df.get("claimed_tenancy", 2.0)

        df["claimed_tenancy"] = pd.to_numeric(df.get("claimed_tenancy", 2.0), errors="coerce").fillna(2.0)
        df["working_tenancy"] = pd.to_numeric(df["working_tenancy"], errors="coerce").fillna(df["claimed_tenancy"])
        df["tenancy_diff"] = df["working_tenancy"] - df["claimed_tenancy"]

        # Calculate Outage Penalties
        penalties = []
        excess_mins_list = []
        outage_notes = []

        for mins in df["outage_minutes"]:
            pen_amt, exc_mins, note = calculate_outage_penalty(mins, self.days_in_month)
            penalties.append(pen_amt)
            excess_mins_list.append(exc_mins)
            outage_notes.append(note)

        df["excess_outage_minutes"] = excess_mins_list
        df["outage_penalty_rs"] = penalties
        df["outage_audit_note"] = outage_notes

        # Calculate Verified Energy Charges if energy parameters are present
        if "dg_run_hours" in df.columns and "dg_cph" in df.columns:
            calc_monthly = []
            for _, row in df.iterrows():
                eng = calculate_energy_charges(
                    dg_run_hours=row.get("dg_run_hours", 1.0),
                    dg_cph=row.get("dg_cph", 0.82),
                    diesel_rate=row.get("diesel_rate", 91.88),
                    eb_rate=row.get("eb_rate", 8.40),
                    eb_cph=row.get("eb_cph", 1.17),
                    days_in_month=self.days_in_month
                )
                calc_monthly.append(eng["total_monthly"])
            df["calculated_energy_amount"] = calc_monthly
        else:
            df["calculated_energy_amount"] = df.get("claimed_amount", 0.0)

        # Baseline amount check
        claimed = pd.to_numeric(df.get("claimed_amount", 0.0), errors="coerce").fillna(0.0)
        calc = pd.to_numeric(df["calculated_energy_amount"], errors="coerce").fillna(claimed)
        
        # BSNL standard policy: Base claim taken is MIN(Claimed, Calculated)
        df["passed_base_amount"] = np.minimum(claimed, calc)

        # Net Payable = Passed Base - Outage Penalty
        df["net_payable_amount"] = np.maximum(0.0, df["passed_base_amount"] - df["outage_penalty_rs"])

        # Discrepancy / Overbilled Amount = Claimed - Net Payable
        df["total_deduction"] = claimed - df["net_payable_amount"]

        # Status Flag
        status = []
        for _, row in df.iterrows():
            ded = row["total_deduction"]
            pen = row["outage_penalty_rs"]
            tdiff = row["tenancy_diff"]
            remark = str(row.get("field_remark", ""))

            if pen > 0 or ded > 50.0 or tdiff != 0:
                status.append("DISPUTED / DEDUCTION")
            elif "theft" in remark.lower() or "down" in remark.lower():
                status.append("FIELD WARNING")
            else:
                status.append("VERIFIED OK")

        df["audit_status"] = status
        return df

    def get_summary_metrics(self, reconciled_df: pd.DataFrame) -> Dict[str, Any]:
        """Generate high-level audit summary metrics."""
        total_sites = len(reconciled_df)
        total_claimed = float(reconciled_df.get("claimed_amount", pd.Series([0])).sum())
        total_passed = float(reconciled_df.get("net_payable_amount", pd.Series([0])).sum())
        total_penalty = float(reconciled_df.get("outage_penalty_rs", pd.Series([0])).sum())
        total_deduction = float(reconciled_df.get("total_deduction", pd.Series([0])).sum())

        disputed_sites = int((reconciled_df["audit_status"] == "DISPUTED / DEDUCTION").sum())
        warning_sites = int((reconciled_df["audit_status"] == "FIELD WARNING").sum())
        verified_sites = int((reconciled_df["audit_status"] == "VERIFIED OK").sum())

        gst_rate = 0.18
        return {
            "total_sites": total_sites,
            "total_claimed_before_tax": round(total_claimed, 2),
            "total_claimed_with_gst": round(total_claimed * (1.0 + gst_rate), 2),
            "total_passed_before_tax": round(total_passed, 2),
            "total_passed_with_gst": round(total_passed * (1.0 + gst_rate), 2),
            "total_outage_penalty": round(total_penalty, 2),
            "total_deduction_before_tax": round(total_deduction, 2),
            "total_deduction_with_gst": round(total_deduction * (1.0 + gst_rate), 2),
            "disputed_sites_count": disputed_sites,
            "warning_sites_count": warning_sites,
            "verified_sites_count": verified_sites
        }
