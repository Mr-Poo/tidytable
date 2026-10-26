"""Reading, cleaning, validating and writing tables."""
from __future__ import annotations

import csv
import datetime as dt
import io
import numbers
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

CURRENCY = "$€£¥₫₹"


class TidyError(Exception):
    """A user-facing error (bad input, bad option, unwritable output)."""


def _title(s: str) -> str:
    return re.sub(r"\w[\w']*", lambda m: m[0][0].upper() + m[0][1:].lower(), s)


STYLES = {"upper": str.upper, "lower": str.lower, "title": _title}


def _num(x: float):
    return int(x) if float(x).is_integer() and abs(x) < 1e15 else x


def _plain(num: str):
    """Parse digits with , and . separators; None when ambiguous/invalid."""
    seps = {c for c in ",." if c in num}
    dec = th = None
    if len(seps) == 2:
        dec = "," if num.rfind(",") > num.rfind(".") else "."
        th = ",.".replace(dec, "")
    elif seps:
        c = seps.pop()
        head, _, tail = num.rpartition(c)
        if num.count(c) > 1:
            th = c
        elif len(tail) == 3 and re.fullmatch(r"[1-9]\d{0,2}", head):
            if c == ".":
                return None  # 1.234 could be 1234 or 1.234: do not guess
            th = c
        else:
            dec = c
    ip, _, fp = num.partition(dec) if dec else (num, "", "")
    if dec and not fp.isdigit():
        return None
    if th:
        if not re.fullmatch(rf"\d{{1,3}}(?:{re.escape(th)}\d{{3}})+", ip):
            return None
        ip = ip.replace(th, "")
    return float(ip + ("." + fp if fp else ""))


def parse_number(v, lenient: bool = False):
    """Return int/float for a number-like value, else None.

    lenient=True also drops stray letters at either end ('12kg', 'USD 7').
    """
    if isinstance(v, (int, float)):
        return v
    t = "".join(c for c in v if c not in CURRENCY and not c.isspace())
    neg = t.startswith("(") and t.endswith(")")
    if neg:
        t = t[1:-1]
    if lenient:
        t = re.fullmatch(r"[^\W\d_]*(.*?)[^\W\d_]*", t)[1]
    m = re.fullmatch(r"([+-]?)(\d[\d.,]*|[.,]\d+)", t)
    if not m or (neg and m[1]):
        return None
    n = _plain(m[2])
    if n is None:
        return None
    return _num(-n if neg or m[1] == "-" else n)


def clean_text(s: str) -> str:
    """Drop non-printing chars, collapse whitespace, trim."""
    return " ".join("".join(c for c in s if c.isprintable() or c.isspace()).split())


def _cell(v):
    if isinstance(v, str):
        return v
    if pd.isna(v):
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, numbers.Integral):
        return int(v)
    if isinstance(v, numbers.Real):
        return _num(float(v))
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if v.time() == dt.time(0) else v.isoformat(sep=" ")
    return str(v)


def read_table(path: Path, sheet: str | None = None):
    """Return (header, rows) from a .csv or .xlsx file."""
    if not path.is_file():
        raise TidyError(f"input file not found: {path}")
    ext = path.suffix.lower()
    if ext not in (".csv", ".xlsx"):
        raise TidyError(f"unsupported input type '{ext}': only .csv and .xlsx")
    try:
        if ext == ".csv":
            raw = path.read_bytes()
            try:
                text = raw.decode("utf-8-sig")
            except UnicodeDecodeError:
                text = raw.decode("cp1252", errors="replace")
            first = text.splitlines()[0] if text.strip() else ""
            sep = max(",;\t|", key=first.count)
            df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str, keep_default_na=False)
        else:
            df = pd.read_excel(path, sheet_name=sheet or 0, dtype=object, keep_default_na=False)
    except Exception as e:  # pandas/openpyxl raise many types
        raise TidyError(f"cannot read {path.name}: {e}") from e
    if df.empty:
        raise TidyError(f"{path.name} has no data rows")
    header = [clean_text(str(c)) for c in df.columns]
    return header, [[_cell(v) for v in r] for r in df.itertuples(index=False)]


def write_table(path: Path, header, rows) -> None:
    """Write a .csv or .xlsx file (formula-looking text stays text)."""
    if path.suffix.lower() == ".csv":
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(rows)
        return
    wb = Workbook()
    ws = wb.active
    ws.append(list(header))
    for r in rows:
        ws.append([None if v == "" else v for v in r])
    for row in ws.iter_rows():
        for c in row:
            if c.data_type == "f":
                c.data_type = "s"
    wb.save(path)


@dataclass
class Options:
    case: list = field(default_factory=list)  # [(style, [cols])]
    key: list = field(default_factory=list)
    required: list = field(default_factory=list)
    no_negative: list = field(default_factory=list)


@dataclass
class Result:
    header: list
    rows: list
    flagged: list  # (row number, action, reason, row values)
    rows_in: int
    dups: int
    fixed: int


def _resolve(header, name, label):
    if name in header:
        return header.index(name)
    hits = [i for i, h in enumerate(header) if h.casefold() == name.casefold()]
    if len(hits) == 1:
        return hits[0]
    raise TidyError(f"unknown column '{name}' in {label} (available: {', '.join(header)})")


def process(header, rows, opts: Options) -> Result:
    """Normalize, repair numbers, validate and dedupe. Mutates rows."""
    cols = lambda names, label: [_resolve(header, n, label) for n in names]  # noqa: E731
    styles = [(STYLES[s], cols(c, "--case")) for s, c in opts.case]
    key, req = cols(opts.key, "--key"), cols(opts.required, "--required")
    neg = cols(opts.no_negative, "--no-negative")
    orig = [[str(v) for v in r] for r in rows]
    bad = [[] for _ in rows]   # reasons that remove the row
    warn = [[] for _ in rows]  # reasons that only flag the row

    for r in rows:
        for j, v in enumerate(r):
            if isinstance(v, str):
                r[j] = clean_text(v)
        for fn, cj in styles:
            for j in cj:
                if isinstance(r[j], str):
                    r[j] = fn(r[j])

    for j in range(len(header)):
        vals = [r[j] for r in rows if r[j] != ""]
        hits = sum(parse_number(v) is not None for v in vals)
        zeros = any(isinstance(v, str) and re.fullmatch(r"0\d+", v) for v in vals)
        if not (j in neg or (vals and hits / len(vals) >= 0.5 and not zeros)):
            continue
        for i, r in enumerate(rows):
            if r[j] == "":
                continue
            n = parse_number(r[j], lenient=True)
            if n is None:
                warn[i].append(f"cannot parse number in '{header[j]}': {r[j]!r}")
            else:
                r[j] = n

    for i, r in enumerate(rows):
        bad[i] += [f"missing required value in '{header[j]}'" for j in req if r[j] == ""]
        bad[i] += [f"negative value in '{header[j]}': {r[j]}" for j in neg
                   if isinstance(r[j], (int, float)) and r[j] < 0]

    seen, drop = set(), set()
    for i, r in enumerate(rows):
        if bad[i]:
            continue
        k = tuple(str(r[j]).casefold() for j in (key or range(len(header))))
        if not any(k):
            continue
        if k in seen:
            drop.add(i)
        else:
            seen.add(k)

    keep = [i for i in range(len(rows)) if not bad[i] and i not in drop]
    fixed = sum(str(v) != o for i in keep for v, o in zip(rows[i], orig[i]))
    flagged = []
    for i, r in enumerate(rows):
        if bad[i]:
            flagged.append((i + 2, "removed", "; ".join(bad[i] + warn[i]), r))
        elif warn[i] and i not in drop:
            flagged.append((i + 2, "kept for review", "; ".join(warn[i]), r))
    return Result(header, [rows[i] for i in keep], flagged, len(rows), len(drop), fixed)
