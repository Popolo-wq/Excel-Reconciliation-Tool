"""Generate realistic sample data for the Excel Reconciliation Tool.

Creates two workbooks under ``sample_data/``:

* ``crm_orders.xlsx``    - 100 reference orders.
* ``bank_payments.xlsx`` - the same orders with intentional discrepancies:
    - 5 orders removed   -> become ``MISSING_IN_TARGET``
    - 8 amounts changed  -> become ``AMOUNT_MISMATCH``
    - 3 new rows added    -> become ``EXTRA_IN_TARGET``
    - 87 left untouched   -> ``MATCHED``

The seed is fixed so the data set is reproducible run-to-run.
"""

from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from faker import Faker

SEED = 42
N_ORDERS = 100
N_MISSING = 5
N_MISMATCH = 8
N_EXTRA = 3
BASE_DATE = date(2026, 6, 23)  # fixed reference so dates are deterministic

OUT_DIR = Path(__file__).parent / "sample_data"


def build_crm_orders(fake: Faker) -> pd.DataFrame:
    """Create the 100-row reference order book."""
    rows = []
    for i in range(N_ORDERS):
        rows.append(
            {
                "order_id": f"ORD-{10001 + i}",
                "customer_name": fake.name(),
                # Amounts span small and large tickets for a realistic spread.
                "amount": round(random.uniform(20.0, 2000.0), 2),
                "order_date": BASE_DATE - timedelta(days=random.randint(0, 120)),
            }
        )
    return pd.DataFrame(rows)


def build_bank_payments(crm: pd.DataFrame, fake: Faker) -> pd.DataFrame:
    """Derive the bank file from the CRM data, injecting known discrepancies."""
    bank = crm.copy()

    # Pick disjoint row sets for each discrepancy type so categories stay clean.
    all_idx = list(bank.index)
    random.shuffle(all_idx)
    missing_idx = all_idx[:N_MISSING]
    mismatch_idx = all_idx[N_MISSING : N_MISSING + N_MISMATCH]

    # AMOUNT_MISMATCH: nudge the amount by a visible, non-trivial delta.
    for idx in mismatch_idx:
        delta = round(random.uniform(10.0, 250.0), 2)
        sign = random.choice((-1, 1))
        bank.at[idx, "amount"] = round(bank.at[idx, "amount"] + sign * delta, 2)

    # MISSING_IN_TARGET: drop these rows entirely from the bank file.
    bank = bank.drop(index=missing_idx)

    # EXTRA_IN_TARGET: payments with no matching CRM order (e.g. refunds/noise).
    extras = pd.DataFrame(
        {
            "order_id": [f"ORD-9000{i + 1}" for i in range(N_EXTRA)],
            "customer_name": [fake.name() for _ in range(N_EXTRA)],
            "amount": [round(random.uniform(20.0, 2000.0), 2) for _ in range(N_EXTRA)],
            "order_date": [
                BASE_DATE - timedelta(days=random.randint(0, 120))
                for _ in range(N_EXTRA)
            ],
        }
    )
    bank = pd.concat([bank, extras], ignore_index=True)

    # Shuffle so the discrepancies are not clustered at predictable positions.
    return bank.sample(frac=1, random_state=SEED).reset_index(drop=True)


def main() -> None:
    """Generate both workbooks and report what was written."""
    random.seed(SEED)
    fake = Faker()
    Faker.seed(SEED)

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    crm = build_crm_orders(fake)
    bank = build_bank_payments(crm, fake)

    crm.to_excel(OUT_DIR / "crm_orders.xlsx", index=False)
    bank.to_excel(OUT_DIR / "bank_payments.xlsx", index=False)

    print(f"Wrote {len(crm)} CRM orders and {len(bank)} bank payments to {OUT_DIR}")


if __name__ == "__main__":
    main()
