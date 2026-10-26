"""Command-line interface for TidyTable."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .core import STYLES, Options, TidyError, process, read_table, write_table


def _list(text: str) -> list[str]:
    return [p.strip() for p in text.split(",") if p.strip()]


def _case(text: str):
    style, _, cols = text.partition(":")
    if style not in STYLES or not _list(cols):
        raise argparse.ArgumentTypeError(
            f"'{text}': expected STYLE:col1,col2 with STYLE in {', '.join(STYLES)}")
    return style, _list(cols)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="tidytable", description="Clean, dedupe and validate a CSV/XLSX table offline.")
    p.add_argument("input", help="input .csv or .xlsx (never modified)")
    p.add_argument("-o", "--output", required=True, help="output .csv or .xlsx")
    p.add_argument("--sheet", help="sheet name for .xlsx input (default: first sheet)")
    p.add_argument("--case", type=_case, action="append", default=[], metavar="STYLE:COLS",
                   help="upper|lower|title for columns, e.g. title:name,city (repeatable)")
    p.add_argument("--key", type=_list, default=[], help="dedupe on these columns (default: all)")
    p.add_argument("--required", type=_list, default=[], help="columns that must not be blank")
    p.add_argument("--no-negative", type=_list, default=[], help="numeric columns that must be >= 0")
    p.add_argument("--report", help="write flagged rows to this .csv file")
    p.add_argument("--version", action="version", version=f"tidytable {__version__}")
    return p


def _same(a: Path, b: Path) -> bool:
    return a.resolve() == b.resolve() or (a.exists() and b.exists() and a.samefile(b))


def run(a) -> int:
    src, out = Path(a.input), Path(a.output)
    rep = Path(a.report) if a.report else None
    if out.suffix.lower() not in (".csv", ".xlsx"):
        raise TidyError("output must end in .csv or .xlsx")
    if rep and rep.suffix.lower() != ".csv":
        raise TidyError("--report must end in .csv")
    if _same(src, out) or (rep and _same(src, rep)):
        raise TidyError("refusing to overwrite the input file; choose a different output/report path")
    if rep and _same(out, rep):
        raise TidyError("--report must differ from the output file")
    header, rows = read_table(src, a.sheet)
    res = process(header, rows, Options(a.case, a.key, a.required, a.no_negative))
    try:
        write_table(out, res.header, res.rows)
        if rep:
            write_table(rep, ["row", "action", "reason"] + res.header,
                        [[n, act, why] + r for n, act, why, r in res.flagged])
    except OSError as e:
        raise TidyError(f"cannot write output: {e}") from e
    print(f"tidytable: {src.name} -> {out}")
    print(f"rows in: {res.rows_in}\nrows out: {len(res.rows)}")
    print(f"duplicates removed: {res.dups}\ncells fixed: {res.fixed}\nrows flagged: {len(res.flagged)}")
    if rep:
        print(f"report: {rep}")
    elif res.flagged:
        print("flagged rows (use --report FILE.csv to save them):")
        for n, act, why, _ in res.flagged[:20]:
            print(f"  row {n} [{act}]: {why}")
    return 0


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return run(args)
    except TidyError as e:
        print(f"tidytable: error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
