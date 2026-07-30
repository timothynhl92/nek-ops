"""The sequence counter -- the single authority for document numbers (§5).

Design commitments, each answering a specific line of the brief:

* **Never guesses.** §5 says the counter is the single authority and must never
  be derived. A key with no entry starts at :data:`CUTOVER_START` only when the
  document date is on or after the cutover month; anything earlier is refused,
  because those numbers belong to the pre-cutover manual sequence.
* **No duplicates.** Reservation takes an exclusive lock, so two concurrent
  runs cannot both take 014.
* **No unexplained gaps.** A number is committed only after the document is on
  disk. If a run dies in between, the reservation is released and the number is
  handed out again -- auditors see neither a hole nor a repeat.
* **Provable.** Every issued reference is appended to an audit log, so the
  sequence can be reconstructed independently of this file.

LIVE ISSUANCE IS NOT WIRED. :func:`reserve` and :func:`commit` are implemented
and tested here, but ``generate_pv.py`` has no code path that calls them; it
uses :func:`peek_next`, which never mutates anything. Wiring them is a separate,
reviewed change.
"""

from __future__ import annotations

import json
import os
import tempfile
from contextlib import contextmanager
from datetime import date, datetime, timezone
from pathlib import Path

try:  # Windows
    import msvcrt
except ImportError:  # pragma: no cover - POSIX fallback
    msvcrt = None
    import fcntl

REPO_ROOT = Path(__file__).resolve().parent.parent
COUNTER_FILE = REPO_ROOT / "counters" / "counters.json"
AUDIT_LOG = REPO_ROOT / "counters" / "issued.log"

# Clean break: numbering restarts at 001 from this month (operator decision,
# 2026-07-30). Documents dated earlier belong to the old manual sequence.
CUTOVER = (2026, 9)
CUTOVER_START = 1

SCHEMA_VERSION = 1


class CounterError(RuntimeError):
    """Raised when a number cannot be issued safely."""


def counter_key(doctype: str, entity: str, year: int) -> str:
    return f"{doctype.upper()}|{entity.upper()}|{year}"


def _before_cutover(doc_date: date) -> bool:
    return (doc_date.year, doc_date.month) < CUTOVER


def assert_on_or_after_cutover(doc_date: date) -> None:
    """Refuse document dates that predate the cutover month."""
    if _before_cutover(doc_date):
        raise CounterError(
            f"document date {doc_date:%Y-%m-%d} predates the "
            f"{CUTOVER[0]}-{CUTOVER[1]:02d} cutover. Numbering restarts at "
            f"{CUTOVER_START:03d} from that month; earlier documents belong to "
            "the previous manual sequence and must not be issued by this tool."
        )


def _load(path: Path) -> dict:
    if not path.is_file():
        return {"version": SCHEMA_VERSION, "counters": {}}
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("version") != SCHEMA_VERSION:
        raise CounterError(
            f"{path.name} is schema version {data.get('version')!r}, "
            f"expected {SCHEMA_VERSION}"
        )
    return data


def _atomic_write(path: Path, data: dict) -> None:
    """Write via a temp file in the same directory, then replace.

    A half-written counter file is worse than none: it would either lose the
    sequence or duplicate a number already printed on paper.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".counters-", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, sort_keys=True)
            fh.write("\n")
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise


@contextmanager
def _locked(path: Path):
    """Hold an exclusive lock for the whole read-modify-write cycle."""
    path.parent.mkdir(parents=True, exist_ok=True)
    lock_path = path.with_suffix(".lock")
    with open(lock_path, "a+b") as handle:
        if msvcrt is not None:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_LOCK, 1)
        else:  # pragma: no cover
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX)
        try:
            yield
        finally:
            if msvcrt is not None:
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:  # pragma: no cover
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def peek_next(
    doctype: str,
    entity: str,
    doc_date: date,
    counter_file: Path | None = None,
) -> int:
    """Return the number that *would* be issued. Mutates nothing.

    This is what dry-run mode uses. Because it does not reserve, two dry runs
    return the same number -- which is correct: neither issued a document.
    """
    assert_on_or_after_cutover(doc_date)
    path = Path(counter_file or COUNTER_FILE)
    data = _load(path)
    entry = data["counters"].get(counter_key(doctype, entity, doc_date.year))
    if entry is None:
        return CUTOVER_START
    return int(entry["last_used"]) + 1


def initialise(
    doctype: str,
    entity: str,
    year: int,
    start_from: int,
    counter_file: Path | None = None,
) -> None:
    """Seed a key explicitly (the ``--init-from`` override).

    Kept for doc types or entities that must continue an existing sequence
    rather than restart at 001.
    """
    if not 1 <= start_from <= 999:
        raise CounterError(f"start {start_from} outside 001-999")
    path = Path(counter_file or COUNTER_FILE)
    key = counter_key(doctype, entity, year)
    with _locked(path):
        data = _load(path)
        if key in data["counters"]:
            raise CounterError(
                f"{key} already exists (last used "
                f"{data['counters'][key]['last_used']}); refusing to reseed"
            )
        data["counters"][key] = {"last_used": start_from - 1, "reserved": None}
        _atomic_write(path, data)


# --------------------------------------------------------------------------
# Live issuance -- implemented, deliberately not called yet. See module docstring.
# --------------------------------------------------------------------------

def reserve(doctype: str, entity: str, doc_date: date, counter_file: Path | None = None) -> int:
    """Take the next number under lock, marking it in flight."""
    assert_on_or_after_cutover(doc_date)
    path = Path(counter_file or COUNTER_FILE)
    key = counter_key(doctype, entity, doc_date.year)
    with _locked(path):
        data = _load(path)
        entry = data["counters"].setdefault(key, {"last_used": CUTOVER_START - 1, "reserved": None})
        if entry.get("reserved"):
            raise CounterError(
                f"{key} already has {entry['reserved']} in flight; a previous "
                "run did not finish. Resolve it before issuing another number."
            )
        nxt = int(entry["last_used"]) + 1
        if nxt > 999:
            raise CounterError(f"{key} exhausted at 999 for the year")
        entry["reserved"] = nxt
        _atomic_write(path, data)
    return nxt


def commit(
    doctype: str,
    entity: str,
    doc_date: date,
    sequence: int,
    reference: str,
    filename: str,
    counter_file: Path | None = None,
) -> None:
    """Confirm a reserved number once the document exists on disk."""
    path = Path(counter_file or COUNTER_FILE)
    key = counter_key(doctype, entity, doc_date.year)
    with _locked(path):
        data = _load(path)
        entry = data["counters"].get(key)
        if not entry or entry.get("reserved") != sequence:
            raise CounterError(f"{key} has no reservation for {sequence}")
        entry["last_used"] = sequence
        entry["reserved"] = None
        _atomic_write(path, data)
    _append_audit(reference, filename)


def release(doctype: str, entity: str, doc_date: date, counter_file: Path | None = None) -> None:
    """Drop an in-flight reservation after a failed run, leaving no gap."""
    path = Path(counter_file or COUNTER_FILE)
    key = counter_key(doctype, entity, doc_date.year)
    with _locked(path):
        data = _load(path)
        entry = data["counters"].get(key)
        if entry:
            entry["reserved"] = None
            _atomic_write(path, data)


def _append_audit(reference: str, filename: str) -> None:
    AUDIT_LOG.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with open(AUDIT_LOG, "a", encoding="utf-8") as fh:
        fh.write(f"{stamp}\t{reference}\t{filename}\n")
