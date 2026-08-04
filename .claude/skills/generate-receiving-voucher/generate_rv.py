"""Generate a Receiving Voucher draft (README §11 item 2).

The internal record of money *received* -- rent, refunds, reimbursements. Not
the tenant-facing document: that is an Official Receipt (`OR`). README §4 is
explicit that the two are never both called "receipt".

DRY-RUN ONLY. This script cannot issue a document. It never increments the
counter, never writes outside ``output/dryrun/``, and never touches a bank.
``--live`` exists solely to fail with an explanation.

The pipeline lives in ``scripts/voucher.py``, shared with the Payment Voucher.
This file exists to name the document type and nothing else.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "scripts"))

from voucher import RECEIVING_VOUCHER, main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main(RECEIVING_VOUCHER, __doc__))
