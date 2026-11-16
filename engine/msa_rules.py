"""
Master Service Agreement (MSA) and SLA Penalty Rules
Based on BSNL Master Service Agreements with IP Vendors (ATC / Elevar / Indus / Ascend)
"""

import calendar
from datetime import datetime
from typing import Dict, Any, Tuple

# Standard MSA SLA parameters
DEFAULT_UPTIME_TARGET = 0.9990  # 99.90% individual site DC power uptime
PENALTY_RATE_PER_HOUR = 500.0   # Rs. 500 per hour or part thereof
PENALTY_CAP_PER_SITE = 10000.0  # Rs. 10,000 maximum penalty per site per month
MINIMUM_PENALTY = 500.0         # Minimum 1 hour charge if outage exceeds allowed threshold

# Standard Rate Cards for Base IP Fee (Table B)
STANDARD_IP_FEES = {
    # (City_Class, BTS_Type, Tenancy) -> Base Fee
    ("Metro", "ID", 1): 32000.0,
    ("Metro", "OD", 1): 30000.0,
    ("Metro", "ID", 2): 29430.0,
    ("Metro", "OD", 2): 28530.0,
    ("Metro", "ID", 3): 29430.0,
    ("Metro", "OD", 3): 28000.0,
    ("Metro", "ID", 4): 25751.0,
    ("Metro", "OD", 4): 24964.0,

    ("Dist", "ID", 1): 17000.0,
    ("Dist", "OD", 1): 15000.0,
    ("Dist", "ID", 2): 17000.0,
    ("Dist", "OD", 2): 15000.0,
    ("Dist", "ID", 3): 15000.0,
    ("Dist", "OD", 3): 13000.0,
    ("Dist", "ID", 4): 15000.0,
    ("Dist", "OD", 4): 13000.0,

    ("Rural", "ID", 1): 17000.0,
    ("Rural", "OD", 1): 15000.0,
    ("Rural", "ID", 2): 17000.0,
    ("Rural", "OD", 2): 15000.0,
    ("Rural", "ID", 3): 15000.0,
    ("Rural", "OD", 3): 13000.0,
    ("Rural", "ID", 4): 15000.0,
    ("Rural", "OD", 4): 13000.0,
}

# 4G Add-on rate structure
RATE_4G_RRU = 800.0  # Rs. 800 per 4G RRU unit
ESCALATION_RATE = 0.125  # 12.5% escalation after 5 years


def get_days_in_month(month: int, year: int) -> int:
    """Return the number of days in a given month and year."""
    return calendar.monthrange(year, month)[1]


def calculate_allowed_downtime_minutes(days_in_month: int, uptime_target: float = DEFAULT_UPTIME_TARGET) -> float:
    """
    Allowed downtime (in minutes) for a site in a given month.
    For 28 days: (1 - 0.999) * 28 * 24 * 60 = 40.32 minutes.
    For 30 days: (1 - 0.999) * 30 * 24 * 60 = 43.20 minutes.
    For 31 days: (1 - 0.999) * 31 * 24 * 60 = 44.64 minutes.
    """
    total_minutes = days_in_month * 24 * 60
    allowed_minutes = (1.0 - uptime_target) * total_minutes
    return round(allowed_minutes, 2)


def calculate_outage_penalty(outage_minutes: float, days_in_month: int = 28) -> Tuple[float, float, str]:
    """
    Calculate Outage Penalty based on MSA Clause 7.0:
    - Outage duration beyond agreed downtime threshold.
    - Rate: Rs. 500 per hour (or part thereof / pro-rata as practiced).
    - Capped at Rs. 10,000 per site per month.

    Returns:
        (penalty_amount, excess_minutes, explanation)
    """
    if outage_minutes is None or outage_minutes <= 0:
        return 0.0, 0.0, "No outage recorded"

    allowed_mins = calculate_allowed_downtime_minutes(days_in_month)
    
    if outage_minutes <= allowed_mins:
        return 0.0, 0.0, f"Outage ({outage_minutes}m) within allowed threshold ({allowed_mins}m)"

    excess_mins = outage_minutes - allowed_mins
    
    # Pro-rata penalty with minimum 1 hr slab
    raw_penalty = (excess_mins / 60.0) * PENALTY_RATE_PER_HOUR
    
    # Enforce minimum 1 hour if excess > 0
    penalty = max(MINIMUM_PENALTY, raw_penalty)
    
    # Cap at maximum Rs. 10,000
    is_capped = penalty > PENALTY_CAP_PER_SITE
    final_penalty = min(penalty, PENALTY_CAP_PER_SITE)

    explanation = (
        f"Exceeded by {excess_mins:.1f} mins ({excess_mins/60.0:.2f} hrs). "
        f"Penalty = ₹{final_penalty:.2f}" + (" [Capped at ₹10,000 max]" if is_capped else "")
    )
    return round(final_penalty, 2), round(excess_mins, 2), explanation


def calculate_energy_charges(
    dg_run_hours: float,
    dg_cph: float,
    diesel_rate: float,
    eb_rate: float,
    eb_cph: float,
    days_in_month: int = 28
) -> Dict[str, float]:
    """
    Calculate verified Energy (EB + DG) charges as per HPC norms.
    """
    dg_run_hours = float(dg_run_hours or 0.0)
    dg_cph = float(dg_cph or 0.0)
    diesel_rate = float(diesel_rate or 91.88)
    eb_rate = float(eb_rate or 8.40)
    eb_cph = float(eb_cph or 0.0)

    # DG charges per day = A * B * C
    dg_per_day = dg_run_hours * dg_cph * diesel_rate

    # EB hours per day = 24 - DG hours
    eb_hours = max(0.0, 24.0 - dg_run_hours)
    
    # EB charges per day = X * Y * Z
    eb_per_day = eb_hours * eb_rate * eb_cph

    total_per_day = dg_per_day + eb_per_day
    total_monthly = total_per_day * days_in_month

    return {
        "dg_per_day": round(dg_per_day, 2),
        "eb_per_day": round(eb_per_day, 2),
        "total_per_day": round(total_per_day, 2),
        "total_monthly": round(total_monthly, 2)
    }
