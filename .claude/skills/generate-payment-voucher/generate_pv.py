"""Generate a Payment Voucher draft (README §8).

DRY-RUN ONLY. This script cannot issue a document. It never increments the
counter, never writes outside ``output/dryrun/``, and never executes, schedules
or releases a payment (§2). Live issuance is a separate, reviewed change;
``--live`` exists solely to fail with an explanation.

The pipeline lives in ``scripts/voucher.py``, shared with the Receiving
Voucher. This file exists to name the document type and nothing else.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "scripts"))

from voucher import PAYMENT_VOUCHER, main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main(PAYMENT_VOUCHER, __doc__))
