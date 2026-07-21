# Home Loan CLP (Construction Linked Plan) Simulator

A single-file, self-contained Python simulator that models a construction-linked home loan month-by-month. It tracks builder demands, bank disbursements, customer contributions, interest accrual, EMIs, prepayments, and outstanding balances from the first day until the loan is fully closed.

---

## Table of Contents
1. [Features](#features)
2. [Prerequisites & Installation](#prerequisites--installation)
3. [How to Run](#how-to-run)
4. [Output Files](#output-files)
5. [Configuration & Field Meanings](#configuration--field-meanings)
    - [Property Configuration](#property-configuration)
    - [Loan Configuration](#loan-configuration)
    - [Extra Self-Funding Pool](#extra-self-funding-pool)
    - [Construction Milestone Schedule](#construction-milestone-schedule)
    - [Interest Rate Changes](#interest-rate-changes)
    - [Prepayments](#prepayments)

---

## Features

- **Construction-Linked Disbursal:** Disburses loan amounts in tranches aligned with your builder's construction milestones and payment plan dates.
- **Flexible Self-Funding:** Allows you to specify upfront customer self-funding (e.g., booking amount) and an extra self-funding pool distributed across subsequent milestones to reduce the overall loan size.
- **Full EMI / Pre-EMI Options:** Supports standard Full EMI (principal + interest from day one) and Pre-EMI (interest-only payment during the construction phase).
- **Extra Prepayment Impact Analysis:** Models recurring monthly prepayments (`PREPAY`) or one-off lump sum prepayments, producing a side-by-side comparison with a baseline scenario (no extra prepayments) to show exactly how much interest and time you save.
- **Visual Analytics:** Generates 5+ charts showcasing loan balance curves, construction progress, cumulative bank disbursements, interest vs. principal splits, and prepayment comparisons.

---

## Prerequisites & Installation

Ensure you have Python 3 installed. The simulator depends on standard data science and plotting libraries.

Install the required packages using `pip`:

```bash
pip install pandas numpy matplotlib openpyxl
```

---

## How to Run

Execute the simulator script directly from your terminal:

```bash
python3 main.py
```

### Example Terminal Output:
```text
PREPAY impact written to: /Users/parthsingh/Desktop/LOAN/output/prepay_impact.txt
Simulation complete. 99 months simulated.
Reports and graphs written to: /Users/parthsingh/Desktop/LOAN/output
```

---

## Output Files

All simulation outputs are saved in the `output/` directory:

| Filename | Type | Description |
| :--- | :--- | :--- |
| **`loan_summary.txt`** | Text | High-level summary of loan stats (total interest/principal paid, closure date, average EMI). |
| **`prepay_impact.txt`** | Text | Comparison of tenure and interest saved with vs. without your configured extra monthly prepayment (`PREPAY`). |
| **`simple_schedule.csv`** | CSV | Beginner-friendly table containing Month, Year, Phase, Loan Disbursed, Interest, Principal, EMI, Prepayment, and Outstanding Balance. |
| **`monthly_schedule.csv`** | CSV | Comprehensive monthly breakdown including demand amount, customer/bank contribution splits, cumulative figures, and builder progress %. |
| **`yearly_summary.csv`** | CSV | Annual aggregation of interest paid, principal paid, disbursements, customer contributions, and ending balance. |
| **`prepay_comparison.png`** | Image | Comparison chart of outstanding balance over time: Baseline vs. With Prepayment. |
| **`loan_balance.png`** | Image | Chart showing the outstanding loan balance trajectory over time. |
| **`interest_vs_principal.png`**| Image | Timeline plot comparing interest charged vs. principal repaid monthly. |
| **`construction_progress.png`** | Image | Cumulative builder demand percentages plotted over milestone dates. |
| **`bank_disbursement.png`** | Image | Curve showing cumulative loan tranches disbursed by the bank. |
| **`cumulative_interest.png`** | Image | Cumulative interest paid over the duration of the loan. |

---

## Configuration & Field Meanings

All configuration options are defined directly at the top of [main.py](file:///Users/parthsingh/Desktop/LOAN/main.py) under the `USER CONFIGURATION` section. 

### Property Configuration

- **`PROPERTY_PRICE`** *(float)*: The base flat/unit cost plus infrastructure charges (excluding GST, TDS, registration, and maintenance). Milestone percentages are computed against this base value.
- **`BUILDER_NAME`** *(str)*: Name of the developer receiving milestone payments (e.g., `"Godrej"`).
- **`MAINTENANCE_AND_SINKING_FUND_CHARGES`** *(float)*: One-time fee paid upon possession. Not included in the loan-linked schedule.
- **`BANK_NAME`** *(str)*: Name of the bank providing the home loan.

### Loan Configuration

- **`CUSTOMER_ONLY_UPTO_PERCENT`** *(float)*: The cumulative percentage of milestones that the customer pays entirely out of pocket. For example, setting to `10.0` covers booking and immediate payment milestones, ensuring the bank loan starts only after 10% is self-paid.
- **`LOAN_AMOUNT`** *(float)*: **Automatically calculated** as `TOTAL_REMAINING_TO_FUND` - `EXTRA_SELF_FUND_POOL`.
- **`INTEREST_RATE`** *(float)*: The initial annual floating/fixed interest rate percentage (e.g., `7.25` for 7.25%).
- **`INTEREST_TYPE`** *(str)*: Either `"Floating"` or `"Fixed"`.
- **`TENURE_YEARS`** *(int)*: Total loan tenure in years (e.g., `20`).
- **`FULL_EMI`** *(bool)*: Set to `True` to pay full EMI (principal + interest) starting from the very first disbursement month.
- **`PRE_EMI`** *(bool)*: Set to `True` to pay interest-only during the construction phase (repaying principal only post-possession).
- **`LOAN_START_DATE`** *(date)*: Calendar start date for the loan modeling.
- **`FIRST_DISBURSEMENT_DATE`** *(date)*: Date when the bank releases the first tranche.

### Extra Self-Funding Pool

If you plan to pay some installments or portions of milestones out-of-pocket to reduce your loan burden:

- **`EXTRA_SELF_FUND_POOL`** *(float)*: Total pool size (in Rs.) you wish to pay directly to the builder during construction. Setting this automatically reduces the required loan size.
- **`EXTRA_SELF_FUND_START_INDEX`** *(int)*: The milestone index (1-based) where the extra self-funding begins (e.g. `4` for Excavation).
- **`EXTRA_SELF_FUND_SHARE_PERCENT`** *(float)*: The percentage of each milestone demand from the start index that you pay out of pocket, drawing from the pool until it is exhausted.

### Construction Milestone Schedule

- **`construction_schedule`** *(List[Dict])*: The list of payment stages. Each stage is a dictionary:
  - `milestone`: Description of the milestone (e.g., `"Foundation Completion"`).
  - `percentage`: The payment percentage demanded at this milestone (sum of all milestone percentages must equal exactly `100`).
  - `date`: Estimate of when this milestone will be reached (`"YYYY-MM-DD"`).

### Interest Rate Changes

- **`interest_changes`** *(List[Dict])*: Allows simulation of fluctuating interest rates over the tenure of the loan. 
  - Example: `[{"date": "2027-01-01", "rate": 7.5}]` will increase the interest rate to 7.5% starting January 2027.

### Prepayments

- **`prepayments`** *(List[Dict])*: Irregular, one-off lump-sum prepayments.
  - Example: `[{"date": "2028-06-15", "amount": 500000}]` (prepays Rs. 5,00,000 in June 2028).
- **`PREPAY`** *(float)*: A fixed extra amount paid **every single month** on top of your regular EMI once the loan is in the Full EMI phase. Set to `0` to disable.
- **`PREPAY_START_DATE`** *(date)*: Custom date to start the recurring prepayments. If `None`, it starts automatically on the first month of the Full EMI phase.
