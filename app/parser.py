from __future__ import annotations

import io
import re
from pathlib import PurePosixPath, PureWindowsPath
from typing import Any
from urllib.parse import unquote, urlparse

import pandas as pd


EXPECTED_COLUMNS = [
    "LinkTitle",
    "Hyperlink",
    "HyperlinkExists",
    "Position",
    "Folder",
    "Filename",
    "Extension",
    "LastModificationDate",
    "LastAccessDate",
    "Owner",
    "docID",
]

COLUMN_ALIASES = {
    "linktitle": "LinkTitle",
    "linkname": "LinkTitle",
    "hyperlink": "Hyperlink",
    "hyperlinkexists": "HyperlinkExists",
    "position": "Position",
    "folder": "Folder",
    "filename": "Filename",
    "extension": "Extension",
    "lastmodificationdate": "LastModificationDate",
    "lastaccessdate": "LastAccessDate",
    "owner": "Owner",
    "docid": "docID",
}


def parse_replace_magic_file(filename: str, content: bytes) -> list[dict[str, Any]]:
    suffix = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    buffer = io.BytesIO(content)

    if suffix == "csv":
        raw_frames = {"CSV": pd.read_csv(buffer, dtype=str, keep_default_na=False, header=None, engine="python")}
    elif suffix == "xlsx":
        raw_frames = pd.read_excel(buffer, dtype=str, keep_default_na=False, engine="openpyxl", sheet_name=None, header=None)
    else:
        raise ValueError("Unsupported file type. Please upload a CSV or XLSX file.")

    df = _find_replace_magic_table(raw_frames)
    df = df.fillna("")
    rows: list[dict[str, Any]] = []

    for index, record in enumerate(df.to_dict(orient="records"), start=1):
        normalized = _normalize_record(record)
        old_path = normalized.get("LinkTitle", "")
        old_url = normalized.get("Hyperlink", "")
        search_file_name = build_search_filename(normalized)

        rows.append(
            {
                "rowNumber": index,
                "raw": normalized,
                "OldPath": old_path,
                "OldURL": old_url,
                "SearchFileName": search_file_name,
            }
        )

    return rows


def _normalize_record(record: dict[str, Any]) -> dict[str, str]:
    normalized: dict[str, str] = {}
    for key, value in record.items():
        clean_key = _clean(key)
        canonical_key = COLUMN_ALIASES.get(_normalize_column_name(clean_key), clean_key)
        normalized[canonical_key] = _clean(value)
    return normalized


def _find_replace_magic_table(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    fallback: pd.DataFrame | None = None

    for sheet_name, frame in frames.items():
        header_index = _find_header_row(frame)
        if header_index is None:
            continue

        headers = _make_unique_headers([_clean(value) for value in frame.iloc[header_index].tolist()])
        table = frame.iloc[header_index + 1 :].copy()
        table.columns = headers
        table = table.loc[:, [column for column in table.columns if column]]
        table = table.dropna(how="all")
        table = table[~table.apply(_is_blank_row, axis=1)]
        if table.empty:
            continue

        normalized_headers = {_normalize_column_name(value) for value in table.columns}
        if {"linktitle", "hyperlink"} <= normalized_headers:
            return table

        if fallback is None:
            fallback = table

    if fallback is not None:
        return fallback

    raise ValueError(
        "Could not find ReplaceMagic columns. Expected a header row containing LinkTitle/Link Title, Hyperlink, or Filename."
    )


def _find_header_row(frame: pd.DataFrame) -> int | None:
    for index, row in frame.iterrows():
        normalized_headers = {_normalize_column_name(value) for value in row.tolist() if _clean(value)}
        if {"linktitle", "hyperlink"} <= normalized_headers:
            return int(index)
        if "filename" in normalized_headers and ("extension" in normalized_headers or "linktitle" in normalized_headers):
            return int(index)
    return None


def _normalize_column_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", _clean(value).casefold())


def _make_unique_headers(headers: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    unique_headers: list[str] = []
    for header in headers:
        if not header:
            unique_headers.append("")
            continue

        count = seen.get(header, 0)
        seen[header] = count + 1
        unique_headers.append(header if count == 0 else f"{header}_{count + 1}")
    return unique_headers


def _is_blank_row(row: pd.Series) -> bool:
    return all(not _clean(value) for value in row.tolist())


def build_search_filename(row: dict[str, str]) -> str:
    filename = _clean(row.get("Filename", ""))
    extension = _clean(row.get("Extension", "")).lstrip(".")

    if filename:
        if _has_extension(filename) or not extension:
            return filename
        return f"{filename}.{extension}"

    return extract_filename(row.get("LinkTitle", "")) or extract_filename(row.get("Hyperlink", ""))


def extract_filename(value: str | None) -> str:
    text = _clean(value)
    if not text:
        return ""

    text = unquote(text).strip().strip('"').strip("'")
    parsed = urlparse(text)
    candidate = parsed.path if parsed.scheme and parsed.netloc else text

    candidate = candidate.replace("\\", "/")
    candidate = re.sub(r"[?#].*$", "", candidate).strip("/")
    if not candidate:
        return ""

    name = PurePosixPath(candidate).name or PureWindowsPath(candidate).name
    return name.strip()


def _has_extension(filename: str) -> bool:
    name = PurePosixPath(filename.replace("\\", "/")).name
    return "." in name and not name.endswith(".")


def _clean(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()
