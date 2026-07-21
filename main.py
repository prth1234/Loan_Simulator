"""
=============================================================================
HOME LOAN CLP (CONSTRUCTION LINKED PLAN) SIMULATOR
=============================================================================

A single-file, self-contained simulator that models a construction-linked
home loan month by month, from loan start until the loan is fully closed.

Run with:
    python home_loan_simulator.py

Outputs (written to ./output/):
    monthly_schedule.csv
    yearly_summary.csv
    loan_summary.txt
    loan_balance.png
    interest_vs_principal.png
    construction_progress.png
    bank_disbursement.png
    cumulative_interest.png

Dependencies: pandas, numpy, matplotlib, openpyxl, datetime, pathlib
=============================================================================
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import List, Dict, Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


##########################
# USER CONFIGURATION
##########################

# -----------------------------------------------------------------------
# NOTE ON DATES: The construction milestone dates below are ESTIMATES,
# built by combining the payment-plan percentages from your agreement's
# "PAYMENT PLAN" table with the tentative completion dates from
# "ANNEXURE-A" (stage-wise schedule) wherever the milestones could be
# reasonably matched. Annexure-A does not name every payment-plan
# milestone individually (e.g. it has no separate line for "Foundation"
# or for each individual floor slab), so those in-between dates have
# been interpolated across the known anchor dates. Please open this file
# and correct any date in `construction_schedule` below to match your
# actual builder demand letters as they arrive -- everything downstream
# recalculates automatically from these dates.
# -----------------------------------------------------------------------

# -----------------------------------------------------------------------
# FIX APPLIED: PROPERTY_PRICE was previously set to the rounded, verbal
# "Rs. 1.45 Cr" figure (14,500,000). All milestone percentages in your
# agreement actually apply to the "Sale Consideration (A)" from Annexure C
# / the cost sheet -- Rs. 1,34,25,243.48 (Flat/Unit Cost + Infra Charges) --
# NOT the rounded headline price. That is confirmed by the Payment Plan
# table itself: 5% of 1,34,25,243.48 = Rs. 6,71,262.17, which matches the
# "On Booking" and "Within 15 Days" rows exactly. GST, TDS, and the
# Advance Maintenance/Sinking Fund charges are NOT part of this base and
# are handled separately (GST/TDS are informational only here; sinking
# fund is tracked in MAINTENANCE_AND_SINKING_FUND_CHARGES below).
# -----------------------------------------------------------------------

# ---- Property ----
PROPERTY_PRICE: float = 13_425_243.48     # Sale Consideration (A): Flat/Unit Cost + Infra Charges — this is the base every milestone % is computed against
BUILDER_NAME: str = "Godrej"              # developer receiving the milestone payments
MAINTENANCE_AND_SINKING_FUND_CHARGES: float = 279_832.32  # one-time charge on possession, NOT part of the 100% collection / loan-linked schedule (informational only)
BANK_NAME: str = "Your Bank"

# ---- Loan ----
# You've already self-paid the first 10% (Booking 5% + Within 15 Days 5% =
# 10% of the Sale Consideration, ~Rs. 13.43L) out of pocket.

# Customer self-funds (pays 100% out of pocket, Rs.0 bank disbursement) every
# milestone whose CUMULATIVE percentage falls at or below this value. The
# loan/bank disbursement only starts from the first milestone that pushes the
# cumulative percentage past this threshold. Set to 0 to disable (loan starts
# from the very first milestone).
CUSTOMER_ONLY_UPTO_PERCENT: float = 10.0  # Booking (5%) + Within 15 Days (5%) = 10% (already self-paid)

# ---- Extra Self-Funded Amount Paid Directly to the Builder ----
# You are taking ONLY the 20-year tenure option, and paying part of certain
# installments directly to Godrej instead of drawing 100% from the bank:
#   - The 3rd milestone ("Within 90 Days") is still paid 100% by the bank.
#   - From the 4th milestone ("Excavation Completion") onward, you pay
#     EXTRA_SELF_FUND_SHARE_PERCENT of each demand directly to the builder,
#     and the bank covers the rest -- until EXTRA_SELF_FUND_POOL is used up.
#   - Once the pool is exhausted, every remaining milestone reverts to
#     100% bank-funded.
#
# LOAN_AMOUNT is DERIVED from this pool, not typed in separately (see the
# formula right below): TOTAL amount still to be funded after your initial
# 10% self-pay = LOAN_AMOUNT + EXTRA_SELF_FUND_POOL, and that TOTAL is held
# constant. So whatever you set EXTRA_SELF_FUND_POOL to, LOAN_AMOUNT
# automatically shrinks or grows to make up the rest -- what you ultimately
# have to pay in total never changes, only the bank/you split does.
EXTRA_SELF_FUND_POOL: float = 1000000.0    # Rs. 10,00,000 — the amount you're covering directly, out of the total still owed
EXTRA_SELF_FUND_START_INDEX: int = 4         # 1-based milestone position where this extra self-funding begins (4 = "Excavation Completion")
EXTRA_SELF_FUND_SHARE_PERCENT: float = 25.0  # % of each demand (from the start index onward) you pay directly to the builder, until the pool runs out

# Total still to be funded after the initial CUSTOMER_ONLY_UPTO_PERCENT
# self-pay -- this is the fixed pie that LOAN_AMOUNT and EXTRA_SELF_FUND_POOL
# always split between them.
TOTAL_REMAINING_TO_FUND: float = PROPERTY_PRICE * (100.0 - CUSTOMER_ONLY_UPTO_PERCENT) / 100.0

# LOAN_AMOUNT = whatever's left of TOTAL_REMAINING_TO_FUND after your
# EXTRA_SELF_FUND_POOL. Change EXTRA_SELF_FUND_POOL above and this updates
# automatically -- you never need to hand-edit LOAN_AMOUNT yourself.
LOAN_AMOUNT: float = TOTAL_REMAINING_TO_FUND - EXTRA_SELF_FUND_POOL
TOTAL_LOAN_REQUIRED: float = LOAN_AMOUNT  # amount actually to be drawn down (capped at this ceiling)
INTEREST_RATE: float = 7.25               # annual %, at loan start
INTEREST_TYPE: str = "Floating"           # "Floating" or "Fixed"
TENURE_YEARS: int = 20                    # total tenure, counted from LOAN_START_DATE
FULL_EMI: bool = True                     # user has opted for FULL EMI (principal+interest) starting from the FIRST disbursement — not interest-only Pre-EMI
PRE_EMI: bool = False                      # Pre-EMI (interest-only during construction) NOT selected
LOAN_START_DATE: date = date(2026, 7, 1)
FIRST_DISBURSEMENT_DATE: date = date(2026, 7, 1)
EMI_START_DATE: Optional[date] = None     # None -> auto-set: FIRST_DISBURSEMENT_DATE if FULL_EMI, else the "Possession" milestone date

# ---- CLP (Construction Linked Plan) ----
# NOTE: superseded by the milestone-index-based logic above
# (CUSTOMER_ONLY_UPTO_PERCENT + EXTRA_SELF_FUND_*). Kept only so
# validate_inputs()'s BANK_SHARE_PERCENT + CUSTOMER_SHARE_PERCENT == 100
# check still passes; not otherwise used in build_construction_schedule().
BANK_SHARE_PERCENT: float = 100.0
CUSTOMER_SHARE_PERCENT: float = 0.0
DISBURSEMENT_MODE: str = "Milestone-linked"
PRE_EMI_UNTIL_POSSESSION: bool = False    # No Pre-EMI phase — pay full EMI (principal + interest) from the very first disbursement
FULL_EMI_AFTER_POSSESSION: bool = True

# Booking amount = customer's share of the first ("Booking") milestone demand.
# Computed automatically further down once the construction schedule is built;
# kept here as a placeholder for documentation purposes only.
BOOKING_AMOUNT: Optional[float] = None

# ---- Construction Schedule ----
# Percentages below are taken directly from the agreement's PAYMENT PLAN
# table (they sum to exactly 100%). Dates are estimates -- see the note above.
construction_schedule: List[Dict] = [
    {"milestone": "Booking",                 "percentage": 5,  "date": "2026-07-01"},
    {"milestone": "Within 15 Days",           "percentage": 5,  "date": "2026-07-15"},
    {"milestone": "Within 90 Days",           "percentage": 10, "date": "2026-09-29"},
    {"milestone": "Excavation Completion",    "percentage": 10, "date": "2027-09-13"},
    {"milestone": "Foundation Completion",    "percentage": 5,  "date": "2027-12-01"},
    {"milestone": "B2 Slab",                  "percentage": 5,  "date": "2028-07-04"},
    {"milestone": "2nd Floor Slab",           "percentage": 5,  "date": "2028-09-01"},
    {"milestone": "5th Floor Slab",           "percentage": 5,  "date": "2028-12-01"},
    {"milestone": "10th Floor Slab",          "percentage": 5,  "date": "2029-03-01"},
    {"milestone": "15th Floor Slab",          "percentage": 5,  "date": "2029-06-01"},
    {"milestone": "20th Floor Slab",          "percentage": 5,  "date": "2029-08-01"},
    {"milestone": "25th Floor Slab",          "percentage": 5,  "date": "2029-10-01"},
    {"milestone": "Flooring",                 "percentage": 10, "date": "2030-06-14"},
    {"milestone": "Occupancy Certificate",    "percentage": 10, "date": "2031-03-01"},
    {"milestone": "Possession",               "percentage": 10, "date": "2031-06-01"},
]

# ---- Interest Rate Changes ----
# List of {"date": "YYYY-MM-DD", "rate": annual_percent}. Leave empty for a
# constant rate for the life of the loan.
interest_changes: List[Dict] = []

# ---- Prepayments ----
# List of {"date": "YYYY-MM-DD", "amount": rupees}. Leave empty for none.
# Use this for one-off, irregular lump-sum prepayments (bonus, gift, etc).
prepayments: List[Dict] = []

# ---- Recurring Monthly Prepayment (PREPAY) ----
# Set PREPAY to a fixed extra amount you want to pay EVERY month, on top of
# your regular EMI, once the loan enters the Full EMI phase. This is separate
# from the one-off `prepayments` list above and repeats automatically every
# month until the loan closes. Set to 0 to disable.
PREPAY: float = 100000.0                 # e.g. Rs. 1,00,000 extra every month
PREPAY_START_DATE: Optional[date] = None  # None -> starts automatically from the EMI start date


##########################
# HELPER FUNCTIONS
##########################

def parse_date(value) -> date:
    """Parse a 'YYYY-MM-DD' string (or pass through a date) into a date object."""
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value))


def days_in_month(year: int, month: int) -> int:
    """Return the number of days in a given month/year."""
    if month == 12:
        nxt = date(year + 1, 1, 1)
    else:
        nxt = date(year, month + 1, 1)
    return (nxt - date(year, month, 1)).days


def add_months(d: date, months: int) -> date:
    """Add a whole number of calendar months to a date, clamping the day."""
    total = d.month - 1 + months
    year = d.year + total // 12
    month = total % 12 + 1
    day = min(d.day, days_in_month(year, month))
    return date(year, month, day)


def months_between(start: date, end: date) -> int:
    """Number of whole calendar months between two dates (end - start)."""
    return (end.year - start.year) * 12 + (end.month - start.month)


##########################
# VALIDATION
##########################

def validate_inputs() -> None:
    """Validate user configuration before running the simulation."""
    total_pct = sum(m["percentage"] for m in construction_schedule)
    if abs(total_pct - 100) > 0.01:
        raise ValueError(
            f"construction_schedule percentages must sum to 100, got {total_pct}"
        )

    if abs(BANK_SHARE_PERCENT + CUSTOMER_SHARE_PERCENT - 100) > 0.01:
        raise ValueError("BANK_SHARE_PERCENT + CUSTOMER_SHARE_PERCENT must equal 100")

    if LOAN_AMOUNT <= 0 or PROPERTY_PRICE <= 0:
        raise ValueError("LOAN_AMOUNT and PROPERTY_PRICE must be positive")

    if LOAN_AMOUNT > PROPERTY_PRICE:
        raise ValueError("LOAN_AMOUNT cannot exceed PROPERTY_PRICE")

    if TENURE_YEARS <= 0:
        raise ValueError("TENURE_YEARS must be positive")

    if INTEREST_RATE <= 0:
        raise ValueError("INTEREST_RATE must be positive")

    dates = [parse_date(m["date"]) for m in construction_schedule]
    if dates != sorted(dates):
        raise ValueError("construction_schedule dates must be in chronological order")

    for ic in interest_changes:
        parse_date(ic["date"])
        if ic["rate"] <= 0:
            raise ValueError("interest_changes rates must be positive")

    for pp in prepayments:
        parse_date(pp["date"])
        if pp["amount"] <= 0:
            raise ValueError("prepayment amounts must be positive")


##########################
# CONSTRUCTION SCHEDULE
##########################

def build_construction_schedule() -> List[Dict]:
    """
    Enrich the raw construction_schedule with computed demand, bank
    contribution, and customer contribution for every milestone.

    Funding rule, applied in order:
      1. Milestones within CUSTOMER_ONLY_UPTO_PERCENT (Booking + Within 15
         Days, 10% total): you self-fund 100%, bank pays Rs. 0.
      2. Milestones after that but before EXTRA_SELF_FUND_START_INDEX (i.e.
         the "Within 90 Days" milestone, #3): bank pays 100%.
      3. From EXTRA_SELF_FUND_START_INDEX onward (#4, "Excavation
         Completion"): you pay EXTRA_SELF_FUND_SHARE_PERCENT of each demand
         directly to the builder, and the bank covers the rest, until your
         EXTRA_SELF_FUND_POOL (Rs. 10L) is fully used up.
      4. Once the extra pool is exhausted, every remaining milestone
         reverts to 100% bank-funded.
    """
    enriched = []
    cumulative_pct = 0.0
    extra_pool_remaining = EXTRA_SELF_FUND_POOL

    for idx, milestone in enumerate(construction_schedule, start=1):
        demand_amount = PROPERTY_PRICE * milestone["percentage"] / 100.0
        cumulative_pct += milestone["percentage"]

        if cumulative_pct <= CUSTOMER_ONLY_UPTO_PERCENT + 1e-9:
            # Rule 1: still within the up-front self-funded portion.
            bank_contribution = 0.0
            customer_contribution = demand_amount
        elif idx < EXTRA_SELF_FUND_START_INDEX:
            # Rule 2: bank-only milestone before the extra self-fund kicks in
            # (e.g. the "Within 90 Days" 3rd milestone).
            bank_contribution = demand_amount
            customer_contribution = 0.0
        elif extra_pool_remaining > 0.01:
            # Rule 3: draw down the extra self-fund pool at the configured
            # share %, capped by whatever remains in the pool.
            customer_contribution = min(
                demand_amount * EXTRA_SELF_FUND_SHARE_PERCENT / 100.0,
                extra_pool_remaining,
            )
            extra_pool_remaining -= customer_contribution
            bank_contribution = demand_amount - customer_contribution
        else:
            # Rule 4: extra self-fund pool exhausted -- back to 100% bank.
            bank_contribution = demand_amount
            customer_contribution = 0.0

        enriched.append({
            "milestone": milestone["milestone"],
            "date": parse_date(milestone["date"]),
            "percentage": milestone["percentage"],
            "demand_amount": demand_amount,
            "bank_contribution": bank_contribution,
            "customer_contribution": customer_contribution,
        })

    # Cap cumulative bank disbursement at TOTAL_LOAN_REQUIRED; any shortfall
    # is pushed onto the customer's contribution for that milestone.
    cumulative_bank = 0.0
    for m in enriched:
        remaining_loan = max(TOTAL_LOAN_REQUIRED - cumulative_bank, 0.0)
        bank_contribution = min(m["bank_contribution"], remaining_loan)
        shortfall = m["bank_contribution"] - bank_contribution
        m["bank_contribution"] = bank_contribution
        m["customer_contribution"] += shortfall
        cumulative_bank += bank_contribution

    return enriched


##########################
# INTEREST RATE LOOKUP
##########################

def get_interest_rate_for_date(as_of: date, schedule_changes: List[Dict]) -> float:
    """Return the applicable annual interest rate (%) for a given date."""
    rate = INTEREST_RATE
    for change in sorted(schedule_changes, key=lambda c: parse_date(c["date"])):
        if parse_date(change["date"]) <= as_of:
            rate = change["rate"]
        else:
            break
    return rate


##########################
# EMI CALCULATION
##########################

def calculate_emi(principal: float, annual_rate_percent: float, remaining_months: int) -> float:
    """
    Standard reducing-balance EMI formula:
        EMI = P * r * (1+r)^n / ((1+r)^n - 1)
    where r is the monthly interest rate and n the remaining number of months.
    """
    if remaining_months <= 0 or principal <= 0:
        return 0.0
    r = annual_rate_percent / 12.0 / 100.0
    if r == 0:
        return principal / remaining_months
    factor = (1 + r) ** remaining_months
    return principal * r * factor / (factor - 1)


##########################
# MAIN MONTHLY SIMULATION
##########################

def simulate_loan(schedule: List[Dict], prepay_override: Optional[float] = None) -> pd.DataFrame:
    """
    Simulate the loan month by month from LOAN_START_DATE until the
    outstanding balance reaches zero. Returns a DataFrame with one row
    per calendar month.

    prepay_override: if given, overrides the global PREPAY amount for this
    run only (used to generate a no-extra-prepayment baseline for comparison
    without touching your actual PREPAY setting).
    """
    monthly_prepay_amount = PREPAY if prepay_override is None else prepay_override
    prepay_start = PREPAY_START_DATE

    if EMI_START_DATE is not None:
        emi_start = EMI_START_DATE
    elif PRE_EMI_UNTIL_POSSESSION:
        emi_start = schedule[-1]["date"]  # Possession milestone
    else:
        emi_start = FIRST_DISBURSEMENT_DATE  # Full EMI from the first disbursement

    rows = []
    current_month = date(LOAN_START_DATE.year, LOAN_START_DATE.month, 1)

    opening_balance = 0.0
    cumulative_disbursed = 0.0
    cumulative_customer_paid = 0.0
    cumulative_bank_paid = 0.0
    total_interest_paid = 0.0
    total_principal_paid = 0.0

    emi_amount = 0.0
    current_emi = 0.0          # the fixed EMI currently in effect, held constant month to month
    recompute_emi = True       # forces an EMI recompute (first month, new disbursement, or rate change)
    last_rate_seen: Optional[float] = None
    max_months = TENURE_YEARS * 12 + 240  # hard safety cap (extra 20 years buffer)
    month_count = 0

    # Total number of EMI months available once full-EMI phase begins,
    # counted from the overall loan tenure (which spans from LOAN_START_DATE).
    total_tenure_months = TENURE_YEARS * 12

    while month_count < max_months:
        month_count += 1
        last_day_of_month = date(current_month.year, current_month.month,
                                  days_in_month(current_month.year, current_month.month))

        rate_today = get_interest_rate_for_date(last_day_of_month, interest_changes)

        # ---- 1) Apply any construction milestones due this month ----
        milestone_names = []
        demand_amount = 0.0
        bank_disbursement = 0.0
        customer_contribution = 0.0
        for m in schedule:
            if m["date"].year == current_month.year and m["date"].month == current_month.month:
                milestone_names.append(m["milestone"])
                demand_amount += m["demand_amount"]
                bank_disbursement += m["bank_contribution"]
                customer_contribution += m["customer_contribution"]

        # `opening_balance` here is the balance carried over from LAST month's
        # closing balance — i.e. it does NOT yet include this month's new
        # disbursement. Interest/EMI for this month is charged on that
        # carried-over balance only. This month's disbursement is added to
        # the balance at the end of the month, so it only starts accruing
        # interest/EMI from NEXT month onwards (matches real-world CLP loans,
        # where EMI on a fresh tranche starts the following month).
        balance_for_this_month = opening_balance

        cumulative_disbursed += bank_disbursement
        cumulative_customer_paid += customer_contribution
        cumulative_bank_paid += bank_disbursement

        # ---- 2) Determine phase: Pre-EMI (interest only) vs Full EMI ----
        in_full_emi_phase = current_month >= date(emi_start.year, emi_start.month, 1)
        phase_label = "Full EMI" if in_full_emi_phase else "Pre-EMI"

        interest_charged = balance_for_this_month * (rate_today / 12.0 / 100.0)

        principal_repaid = 0.0
        emi_this_month = 0.0

        pre_emi_interest_payment = 0.0
        actual_emi = 0.0

        if balance_for_this_month <= 0.01:
            # Nothing disbursed yet (or disbursed only this month, which
            # doesn't start accruing until next month)
            emi_amount = 0.0
        elif not in_full_emi_phase and PRE_EMI_UNTIL_POSSESSION:
            # Pre-EMI: this is NOT an EMI. It's an interest-only payment on
            # whatever has been disbursed so far. No principal is repaid,
            # and the loan balance does not reduce.
            pre_emi_interest_payment = interest_charged
            emi_this_month = 0.0
            principal_repaid = 0.0
        else:
            # Full EMI phase: the EMI is held FIXED month to month (like a
            # real bank EMI) and only recomputed when something changes the
            # math — a new disbursement landing, an interest rate change, or
            # the very first month of this phase. This is what lets PREPAY
            # actually shorten the loan tenure: extra principal paid down
            # between recomputes just brings the closing balance to zero
            # sooner, instead of silently lowering next month's EMI.
            rate_changed = last_rate_seen is not None and rate_today != last_rate_seen
            if recompute_emi or rate_changed:
                months_elapsed_since_start = months_between(
                    date(LOAN_START_DATE.year, LOAN_START_DATE.month, 1), current_month
                )
                remaining_months = max(total_tenure_months - months_elapsed_since_start, 1)
                current_emi = calculate_emi(balance_for_this_month, rate_today, remaining_months)
                recompute_emi = False

            emi_amount = current_emi
            actual_emi = current_emi
            emi_this_month = current_emi
            principal_repaid = max(emi_this_month - interest_charged, 0.0)
            # Guard against overpaying the final month
            principal_repaid = min(principal_repaid, balance_for_this_month)

        last_rate_seen = rate_today
        # A disbursement landing this month raises the balance next month,
        # so the EMI must be recomputed on the following iteration.
        if bank_disbursement > 0.01:
            recompute_emi = True

        # ---- 3) Apply prepayments due this month ----
        # 3a) One-off lump-sum prepayments from the `prepayments` list.
        one_off_prepayment = 0.0
        for pp in prepayments:
            pp_date = parse_date(pp["date"])
            if pp_date.year == current_month.year and pp_date.month == current_month.month:
                one_off_prepayment += pp["amount"]

        # 3b) Recurring monthly PREPAY: a fixed extra amount every month,
        # starting from PREPAY_START_DATE (or the EMI start date if not set),
        # applied only once the loan has an outstanding balance to reduce.
        recurring_prepayment = 0.0
        if monthly_prepay_amount > 0 and balance_for_this_month > 0.01:
            effective_prepay_start = prepay_start if prepay_start is not None else emi_start
            if current_month >= date(effective_prepay_start.year, effective_prepay_start.month, 1):
                recurring_prepayment = monthly_prepay_amount

        prepayment_amount = one_off_prepayment + recurring_prepayment
        prepayment_amount = min(prepayment_amount, max(balance_for_this_month - principal_repaid, 0.0))

        # This month's new disbursement is added AFTER this month's
        # interest/EMI is computed, so it starts accruing next month.
        closing_balance = balance_for_this_month - principal_repaid - prepayment_amount + bank_disbursement
        closing_balance = max(closing_balance, 0.0)

        total_interest_paid += interest_charged
        total_principal_paid += principal_repaid + prepayment_amount

        pct_paid_to_builder = (cumulative_disbursed + cumulative_customer_paid) / PROPERTY_PRICE * 100.0

        rows.append({
            "Month": current_month.strftime("%b"),
            "Year": current_month.year,
            "Date": current_month.isoformat(),
            "Phase": phase_label,
            "Construction Milestone": ", ".join(milestone_names) if milestone_names else "",
            "Demand %": sum(m["percentage"] for m in schedule

                             if m["date"].year == current_month.year
                             and m["date"].month == current_month.month),
            "Demand Amount": round(demand_amount, 2),
            "Bank Disbursement": round(bank_disbursement, 2),
            "Customer Contribution": round(customer_contribution, 2),
            "Total Loan Disbursed": round(cumulative_disbursed, 2),
            "Opening Loan Balance": round(balance_for_this_month, 2),
            "Interest Rate": rate_today,
            "Interest Charged": round(interest_charged, 2),
            "Principal Repaid": round(principal_repaid, 2),
            "Pre-EMI Interest Payment": round(pre_emi_interest_payment, 2),
            "EMI": round(actual_emi, 2),
            "Monthly Payment (Pre-EMI Interest or EMI)": round(emi_this_month, 2),
            "Prepayment": round(prepayment_amount, 2),
            "Closing Balance": round(closing_balance, 2),
            "Total Interest Paid": round(total_interest_paid, 2),
            "Total Principal Paid": round(total_principal_paid, 2),
            "Total Customer Paid": round(cumulative_customer_paid, 2),
            "Total Bank Paid": round(cumulative_bank_paid, 2),
            f"% of Property Value Paid to {BUILDER_NAME}": round(pct_paid_to_builder, 2),
        })

        # ---- 4) Advance to next month ----
        opening_balance = closing_balance
        current_month = add_months(current_month, 1)

        # Stop once the loan is fully disbursed AND fully repaid.
        all_disbursed = current_month > schedule[-1]["date"]
        if all_disbursed and closing_balance <= 0.01:
            break

    return pd.DataFrame(rows)


##########################
# REPORTS
##########################

def write_monthly_schedule(df: pd.DataFrame, output_dir: Path) -> None:
    """Write the full month-by-month schedule to monthly_schedule.csv."""
    df.to_csv(output_dir / "monthly_schedule.csv", index=False)


def write_simple_schedule(df: pd.DataFrame, output_dir: Path) -> None:
    """
    Write a beginner-friendly CSV with just: Month, Year, Phase, Interest,
    Principal, EMI, Prepayment, and Outstanding Balance. EMI = Interest +
    Principal (during Pre-EMI, Principal is 0 so EMI = Interest only).
    """
    builder_pct_col = f"% of Property Value Paid to {BUILDER_NAME}"
    simple = pd.DataFrame({
        "Month": df["Month"],
        "Year": df["Year"],
        "Phase": df["Phase"],
        "Loan Disbursed This Month (Rs.)": df["Bank Disbursement"],
        "Total Loan Disbursed (Rs.)": df["Total Loan Disbursed"],
        "Interest (Rs.)": df["Interest Charged"],
        "Principal (Rs.)": df["Principal Repaid"],
        "Pre-EMI Interest Payment (Rs.)": df["Pre-EMI Interest Payment"],
        "EMI (Rs.)": df["EMI"],
        "Prepayment (Rs.)": df["Prepayment"],
        "Outstanding Loan Balance (Rs.)": df["Closing Balance"],
        builder_pct_col: df[builder_pct_col],
    })
    simple.to_csv(output_dir / "simple_schedule.csv", index=False)



def write_yearly_summary(df: pd.DataFrame, output_dir: Path) -> pd.DataFrame:
    """Aggregate the monthly schedule into a yearly summary CSV."""
    yearly = df.groupby("Year").agg(
        **{
            "Interest Paid": ("Interest Charged", "sum"),
            "Principal Paid": ("Principal Repaid", "sum"),
            "Prepayments": ("Prepayment", "sum"),
            "Outstanding Loan": ("Closing Balance", "last"),
            "Bank Disbursement": ("Bank Disbursement", "sum"),
            "Customer Contribution": ("Customer Contribution", "sum"),
        }
    ).reset_index()
    yearly["Principal Paid"] = yearly["Principal Paid"] + yearly["Prepayments"]
    yearly = yearly[[
        "Year", "Interest Paid", "Principal Paid", "Outstanding Loan",
        "Bank Disbursement", "Customer Contribution",
    ]]
    yearly.to_csv(output_dir / "yearly_summary.csv", index=False)
    return yearly


def summarize_closure(df: pd.DataFrame) -> Dict:
    """Extract closure date, duration, and totals from a simulated DataFrame.
    Reused by both the main loan summary and the PREPAY impact comparison."""
    closure_date = df.loc[df["Closing Balance"] <= 0.01, "Date"]
    closure_date_str = closure_date.iloc[-1] if not closure_date.empty else df["Date"].iloc[-1]
    effective_months = months_between(LOAN_START_DATE, date.fromisoformat(closure_date_str)) + 1
    return {
        "closure_date": closure_date_str,
        "effective_months": effective_months,
        "total_interest": df["Interest Charged"].sum(),
        "total_principal": df["Principal Repaid"].sum() + df["Prepayment"].sum(),
        "total_prepaid": df["Prepayment"].sum(),
    }


def write_loan_summary(df: pd.DataFrame, schedule: List[Dict], output_dir: Path) -> None:
    """Write a plain-text loan summary to loan_summary.txt."""
    closure = summarize_closure(df)
    total_interest = closure["total_interest"]
    total_principal = closure["total_principal"]
    total_customer = df["Customer Contribution"].sum()
    total_bank = df["Bank Disbursement"].sum()
    closure_date_str = closure["closure_date"]
    effective_months = closure["effective_months"]
    emi_rows = df[df["EMI"] > 0]
    average_emi = emi_rows["EMI"].mean() if not emi_rows.empty else 0.0

    lines = [
        "=" * 60,
        "HOME LOAN CLP SIMULATION SUMMARY",
        "=" * 60,
        f"Bank Name                 : {BANK_NAME}",
        f"Builder Name               : {BUILDER_NAME}",
        f"Property Price (base)     : Rs. {PROPERTY_PRICE:,.2f}",
        f"Maintenance/Sinking Fund  : Rs. {MAINTENANCE_AND_SINKING_FUND_CHARGES:,.2f} "
        f"(one-time, on possession, not loan-linked)",
        f"Loan Amount               : Rs. {LOAN_AMOUNT:,.2f}",
        f"Interest Type             : {INTEREST_TYPE}",
        f"Starting Interest Rate    : {INTEREST_RATE:.2f}%",
        f"Tenure (years)            : {TENURE_YEARS}",
        "-" * 60,
        f"Total Interest Paid       : Rs. {total_interest:,.2f}",
        f"Total Principal Paid      : Rs. {total_principal:,.2f}",
        f"Total Customer Contribution: Rs. {total_customer:,.2f}",
        f"Total Bank Contribution    : Rs. {total_bank:,.2f}",
        "-" * 60,
        f"Loan Closure Date         : {closure_date_str}",
        f"Effective Loan Duration   : {effective_months} months "
        f"({effective_months / 12:.1f} years)",
        f"Average EMI (post start)  : Rs. {average_emi:,.2f}",
        "=" * 60,
    ]
    (output_dir / "loan_summary.txt").write_text("\n".join(lines), encoding="utf-8")


def write_prepay_impact(df_with_prepay: pd.DataFrame, df_baseline: pd.DataFrame, output_dir: Path) -> None:
    """
    Write prepay_impact.txt: a side-by-side comparison of the loan WITH your
    configured PREPAY amount vs a baseline with no extra monthly prepayment,
    so you can see exactly how much faster the loan finishes and how much
    interest you save.
    """
    with_pp = summarize_closure(df_with_prepay)
    baseline = summarize_closure(df_baseline)

    months_saved = baseline["effective_months"] - with_pp["effective_months"]
    interest_saved = baseline["total_interest"] - with_pp["total_interest"]

    lines = [
        "=" * 60,
        "PREPAY IMPACT: EXTRA MONTHLY PREPAYMENT COMPARISON",
        "=" * 60,
        f"Extra Monthly Prepayment (PREPAY) : Rs. {PREPAY:,.2f}",
        "-" * 60,
        "WITHOUT extra PREPAY (baseline):",
        f"  Loan Closure Date               : {baseline['closure_date']}",
        f"  Effective Loan Duration         : {baseline['effective_months']} months "
        f"({baseline['effective_months'] / 12:.1f} years)",
        f"  Total Interest Paid             : Rs. {baseline['total_interest']:,.2f}",
        "-" * 60,
        "WITH extra PREPAY:",
        f"  Loan Closure Date               : {with_pp['closure_date']}",
        f"  Effective Loan Duration         : {with_pp['effective_months']} months "
        f"({with_pp['effective_months'] / 12:.1f} years)",
        f"  Total Interest Paid             : Rs. {with_pp['total_interest']:,.2f}",
        f"  Total Extra Prepaid             : Rs. {with_pp['total_prepaid']:,.2f}",
        "=" * 60,
        "IMPACT:",
        f"  Loan Finishes Sooner By         : {months_saved} months "
        f"({months_saved / 12:.1f} years)",
        f"  Interest Saved                  : Rs. {interest_saved:,.2f}",
        "=" * 60,
    ]
    (output_dir / "prepay_impact.txt").write_text("\n".join(lines), encoding="utf-8")


def plot_prepay_comparison(df_with_prepay: pd.DataFrame, df_baseline: pd.DataFrame, output_dir: Path) -> None:
    """Plot outstanding balance WITH vs WITHOUT the extra PREPAY, side by side."""
    plt.figure(figsize=(10, 5))
    plt.plot(pd.to_datetime(df_baseline["Date"]), df_baseline["Closing Balance"],
              label="Without extra PREPAY", linestyle="--")
    plt.plot(pd.to_datetime(df_with_prepay["Date"]), df_with_prepay["Closing Balance"],
              label=f"With PREPAY (Rs. {PREPAY:,.0f}/month)")
    plt.title("Outstanding Loan Balance: With vs Without Extra Prepayment")
    plt.xlabel("Date")
    plt.ylabel("Balance (Rs.)")
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(output_dir / "prepay_comparison.png", dpi=150)
    plt.close()


##########################
# GRAPHS
##########################

def plot_loan_balance(df: pd.DataFrame, output_dir: Path) -> None:
    plt.figure(figsize=(10, 5))
    plt.plot(pd.to_datetime(df["Date"]), df["Closing Balance"])
    plt.title("Outstanding Loan Balance Over Time")
    plt.xlabel("Date")
    plt.ylabel("Balance (Rs.)")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(output_dir / "loan_balance.png", dpi=150)
    plt.close()


def plot_interest_vs_principal(df: pd.DataFrame, output_dir: Path) -> None:
    plt.figure(figsize=(10, 5))
    dates = pd.to_datetime(df["Date"])
    plt.plot(dates, df["Interest Charged"], label="Interest Charged")
    plt.plot(dates, df["Principal Repaid"] + df["Prepayment"], label="Principal Repaid")
    plt.title("Monthly Interest vs Principal")
    plt.xlabel("Date")
    plt.ylabel("Amount (Rs.)")
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(output_dir / "interest_vs_principal.png", dpi=150)
    plt.close()


def plot_construction_progress(schedule: List[Dict], output_dir: Path) -> None:
    dates = [m["date"] for m in schedule]
    cumulative_pct = np.cumsum([m["percentage"] for m in schedule])
    plt.figure(figsize=(10, 5))
    plt.step(dates, cumulative_pct, where="post", marker="o")
    plt.title("Construction Progress (Cumulative % Demanded)")
    plt.xlabel("Date")
    plt.ylabel("Cumulative %")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(output_dir / "construction_progress.png", dpi=150)
    plt.close()


def plot_bank_disbursement(df: pd.DataFrame, output_dir: Path) -> None:
    plt.figure(figsize=(10, 5))
    plt.plot(pd.to_datetime(df["Date"]), df["Total Loan Disbursed"])
    plt.title("Cumulative Bank Disbursement")
    plt.xlabel("Date")
    plt.ylabel("Amount Disbursed (Rs.)")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(output_dir / "bank_disbursement.png", dpi=150)
    plt.close()


def plot_cumulative_interest(df: pd.DataFrame, output_dir: Path) -> None:
    plt.figure(figsize=(10, 5))
    plt.plot(pd.to_datetime(df["Date"]), df["Total Interest Paid"])
    plt.title("Cumulative Interest Paid")
    plt.xlabel("Date")
    plt.ylabel("Amount (Rs.)")
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(output_dir / "cumulative_interest.png", dpi=150)
    plt.close()


##########################
# MAIN
##########################

def main() -> None:
    """Entry point: validate inputs, run the simulation, write all reports."""
    validate_inputs()

    output_dir = Path(__file__).resolve().parent / "output"
    output_dir.mkdir(exist_ok=True)

    schedule = build_construction_schedule()
    df = simulate_loan(schedule)

    write_monthly_schedule(df, output_dir)
    write_simple_schedule(df, output_dir)
    write_yearly_summary(df, output_dir)
    write_loan_summary(df, schedule, output_dir)

    plot_loan_balance(df, output_dir)
    plot_interest_vs_principal(df, output_dir)
    plot_construction_progress(schedule, output_dir)
    plot_bank_disbursement(df, output_dir)
    plot_cumulative_interest(df, output_dir)

    # If a recurring PREPAY amount is configured, also run a no-extra-prepay
    # baseline and report exactly how much sooner the loan closes and how
    # much interest is saved.
    if PREPAY > 0:
        df_baseline = simulate_loan(schedule, prepay_override=0.0)
        write_prepay_impact(df, df_baseline, output_dir)
        plot_prepay_comparison(df, df_baseline, output_dir)
        print(f"PREPAY impact written to: {output_dir / 'prepay_impact.txt'}")

    print(f"Simulation complete. {len(df)} months simulated.")
    print(f"Reports and graphs written to: {output_dir}")


if __name__ == "__main__":
    main()