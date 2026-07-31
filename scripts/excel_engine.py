"""Excel COM engine: recalculation and PDF export for the generation layer.

Per README §7, openpyxl writes formulas as strings with no cached values, so a
filled template must be recalculated by a real spreadsheet engine before its
totals can be read back or printed. The brief assumed LibreOffice; this machine
has Microsoft 365, which recalculates with Excel's own engine and exports PDF
with exact fidelity to the template's print setup.

Every entry point drives a *private* Excel instance (``DispatchEx``), so a
workbook the operator happens to have open is never touched, and quitting here
never closes their session.
"""

from __future__ import annotations

import contextlib
from collections.abc import Iterator
from pathlib import Path

import pythoncom
import win32com.client

# Excel constants, hard-coded so we do not depend on a generated COM cache.
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_CALCULATION_AUTOMATIC = -4105

# Printers whose drivers are local and always answer immediately.
LOCAL_PRINTERS = ("Microsoft Print to PDF", "Microsoft XPS Document Writer")


class RecalcError(RuntimeError):
    """Raised when Excel cannot open or evaluate a workbook."""


@contextlib.contextmanager
def excel_app() -> Iterator["win32com.client.CDispatch"]:
    """Yield a hidden, private Excel application, always torn down."""
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.AskToUpdateLinks = False
        app.AlertBeforeOverwriting = False
        yield app
    finally:
        if app is not None:
            with contextlib.suppress(Exception):
                app.Quit()
        del app
        pythoncom.CoUninitialize()


@contextlib.contextmanager
def open_workbook(app, path: Path, read_only: bool = False) -> Iterator["win32com.client.CDispatch"]:
    """Open ``path`` in ``app``; close without saving unless told otherwise."""
    if not path.is_file():
        raise FileNotFoundError(path)
    # Excel's COM API resolves relative paths against its own working
    # directory, not ours, so an absolute path is mandatory here.
    wb = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=read_only)

    # If the file is already open elsewhere, Excel hands back a read-only copy
    # instead of refusing. With DisplayAlerts off, a later Save() is then a
    # silent no-op: the run reports success and nothing reaches disk. Fail here
    # instead, while the cause is still obvious.
    if not read_only and wb.ReadOnly:
        with contextlib.suppress(Exception):
            wb.Close(SaveChanges=False)
        raise RecalcError(
            f"{path.name} opened read-only, so changes could not be saved. "
            "It is almost certainly open in another Excel window -- close it "
            "and run again."
        )

    # Excel rejects the Calculation property until a workbook exists, and a
    # template saved in manual mode would otherwise silently skip recalc.
    with contextlib.suppress(Exception):
        app.Calculation = XL_CALCULATION_AUTOMATIC
    try:
        yield wb
    finally:
        with contextlib.suppress(Exception):
            wb.Close(SaveChanges=False)


def _use_local_printer(app, preferred: str | None = None) -> str | None:
    """Point Excel at a local printer; return the previous ActivePrinter.

    ``ExportAsFixedFormat`` lays out pages through the *active printer's*
    driver. This machine's Windows default is a Brother on a WSD port, which
    blocks indefinitely when the device is asleep while still reporting its
    status as "Normal" — the export never returns and Excel must be killed.
    Against a local driver the same export completes in well under a second.

    Excel addresses printers as ``"<name> on <port>"`` using its own ``NeNN:``
    aliases, which do not match the Windows port names, so the alias is found
    by trying them.
    """
    try:
        original = app.ActivePrinter
    except Exception:  # noqa: BLE001 - property is unavailable in some states
        original = None

    candidates = [preferred] if preferred else list(LOCAL_PRINTERS)
    for name in filter(None, candidates):
        if " on " in name:  # caller supplied a fully-qualified device string
            try:
                app.ActivePrinter = name
                return original
            except Exception:  # noqa: BLE001 - fall through to alias search
                pass
        for index in range(20):
            try:
                app.ActivePrinter = f"{name} on Ne{index:02d}:"
                return original
            except Exception:  # noqa: BLE001 - wrong alias, keep looking
                continue
    return original


def export_worksheet(
    app,
    wb,
    sheet: str,
    dest: str | Path,
    printer: str | None = None,
    draft_header: str | None = None,
) -> Path:
    """Export one worksheet of an already-open workbook to PDF.

    Exposed separately from :func:`export_pdf` so a caller can sync, fill,
    recalculate, verify and export inside a single Excel session. Launching
    Excel costs roughly fifty seconds, so doing that once instead of five times
    is the difference between a usable tool and an unusable one.

    ``draft_header`` writes a centre page header on the worksheet. It is set on
    the caller's working copy, never on the template on disk.
    """
    dest = Path(dest).resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)
    ws = wb.Worksheets(sheet)

    # Switch printers *before* touching PageSetup: writing any PageSetup
    # property round-trips to the active printer's driver, so setting a header
    # while the WSD default is selected stalls exactly as the export does.
    original_printer = _use_local_printer(app, printer)
    if draft_header is not None:
        ws.PageSetup.CenterHeader = draft_header

    try:
        ws.ExportAsFixedFormat(
            Type=XL_TYPE_PDF,
            Filename=str(dest),
            Quality=XL_QUALITY_STANDARD,
            IncludeDocProperties=False,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
    except Exception as exc:  # noqa: BLE001 - surfaced with context below
        raise RecalcError(f"PDF export failed for sheet {sheet!r}: {exc}") from exc
    finally:
        if original_printer:
            with contextlib.suppress(Exception):
                app.ActivePrinter = original_printer

    if not dest.is_file():
        raise RecalcError(f"Excel reported success but {dest} was not written")
    return dest


def sheet_names(path: str | Path) -> list[str]:
    """Return the worksheet names of ``path``, in tab order."""
    src = Path(path).resolve()
    with excel_app() as app, open_workbook(app, src, read_only=True) as wb:
        return [ws.Name for ws in wb.Worksheets]


def recalc(path: str | Path) -> Path:
    """Fully rebuild every formula in ``path`` and save it in place."""
    target = Path(path).resolve()
    with excel_app() as app, open_workbook(app, target) as wb:
        app.CalculateFullRebuild()
        try:
            wb.Save()
        except Exception as exc:  # noqa: BLE001 - surfaced with context below
            raise RecalcError(f"Excel could not save {target.name}: {exc}") from exc
    return target


def read_cells(path: str | Path, sheet: str, cells: dict[str, str]) -> dict[str, object]:
    """Read ``{label: A1_ref}`` from ``sheet`` after a full rebuild.

    Use this for the sanity checks the brief requires — a clean recalc proves
    formulas *evaluate*, not that they are *right*.
    """
    src = Path(path).resolve()
    with excel_app() as app, open_workbook(app, src, read_only=True) as wb:
        app.CalculateFullRebuild()
        ws = wb.Worksheets(sheet)
        return {label: ws.Range(ref).Value for label, ref in cells.items()}


def export_pdf(
    src: str | Path,
    dest: str | Path,
    sheet: str | None = None,
    recalculate: bool = True,
    printer: str | None = None,
) -> Path:
    """Export ``src`` to PDF at ``dest``.

    ``sheet`` exports that worksheet's print area alone; omit it to export the
    whole workbook. The templates keep their Control Panel in J:K outside the
    print area, so a named sheet is what callers normally want.

    ``printer`` overrides the local-printer autodetection; see
    :func:`_use_local_printer` for why the active printer matters here.
    """
    src = Path(src).resolve()
    dest = Path(dest).resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)

    with excel_app() as app, open_workbook(app, src, read_only=not recalculate) as wb:
        if recalculate:
            app.CalculateFullRebuild()
        original_printer = _use_local_printer(app, printer)
        target = wb.Worksheets(sheet) if sheet else wb
        try:
            target.ExportAsFixedFormat(
                Type=XL_TYPE_PDF,
                Filename=str(dest),
                Quality=XL_QUALITY_STANDARD,
                IncludeDocProperties=False,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        except Exception as exc:  # noqa: BLE001 - surfaced with context below
            raise RecalcError(f"PDF export failed for {src.name}: {exc}") from exc
        finally:
            # Excel remembers ActivePrinter per user, so leave it as we found it.
            if original_printer:
                with contextlib.suppress(Exception):
                    app.ActivePrinter = original_printer

    if not dest.is_file():
        raise RecalcError(f"Excel reported success but {dest} was not written")
    return dest


def _main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p_sheets = sub.add_parser("sheets", help="list worksheet names")
    p_sheets.add_argument("path", type=Path)

    p_recalc = sub.add_parser("recalc", help="rebuild formulas and save in place")
    p_recalc.add_argument("path", type=Path)

    p_pdf = sub.add_parser("pdf", help="export to PDF")
    p_pdf.add_argument("src", type=Path)
    p_pdf.add_argument("dest", type=Path)
    p_pdf.add_argument("--sheet", default=None)
    p_pdf.add_argument("--no-recalc", dest="recalculate", action="store_false")
    p_pdf.add_argument("--printer", default=None, help="override printer autodetection")

    args = parser.parse_args()

    if args.command == "sheets":
        for name in sheet_names(args.path):
            print(name)
    elif args.command == "recalc":
        print(recalc(args.path))
    elif args.command == "pdf":
        print(export_pdf(args.src, args.dest, args.sheet, args.recalculate, args.printer))
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
