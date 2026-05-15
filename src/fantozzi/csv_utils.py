from __future__ import annotations

import csv


def sniff_delimiter(sample: str, *, delimiters: str = ";,", default: str = ";") -> str:
    if not sample:
        return default
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=delimiters)
    except csv.Error:
        return _delimiter_from_header(sample, delimiters=delimiters, default=default)
    delimiter = getattr(dialect, "delimiter", None)
    if delimiter in delimiters:
        return delimiter
    return _delimiter_from_header(sample, delimiters=delimiters, default=default)


def _delimiter_from_header(sample: str, *, delimiters: str, default: str) -> str:
    header = sample.splitlines()[0] if sample.splitlines() else ""
    counts = {delimiter: header.count(delimiter) for delimiter in delimiters}
    best = max(counts, key=counts.get, default=default)
    return best if counts.get(best, 0) else default
