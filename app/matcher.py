from __future__ import annotations

from collections import defaultdict
from typing import Protocol


OUTPUT_COLUMNS = [
    "OldPath",
    "OldURL",
    "SearchFileName",
    "MatchedFileName",
    "NewURL",
    "Status",
    "Remark",
]


class SearchClient(Protocol):
    def search_site_files(self, filename: str) -> list[dict[str, str]]:
        ...


def run_exact_matching(rows: list[dict], graph_client: SearchClient) -> list[dict[str, str]]:
    results: list[dict[str, str]] = []
    cache: dict[str, list[dict[str, str]]] = {}

    for row in rows:
        old_path = row.get("OldPath", "")
        old_url = row.get("OldURL", "")
        search_file_name = row.get("SearchFileName", "")
        result = _base_result(old_path, old_url, search_file_name)

        try:
            if not search_file_name:
                result.update({"Status": "ERROR", "Remark": "Could not determine filename from row"})
                results.append(result)
                continue

            cache_key = search_file_name.casefold()
            if cache_key not in cache:
                cache[cache_key] = graph_client.search_site_files(search_file_name)

            exact_matches = [
                item
                for item in cache[cache_key]
                if item.get("name", "").casefold() == search_file_name.casefold()
            ]
            result.update(_status_from_matches(search_file_name, exact_matches))
        except Exception as exc:
            result.update({"Status": "ERROR", "Remark": str(exc)})

        results.append(result)

    return results


def summarize_results(results: list[dict[str, str]]) -> dict[str, int]:
    counts = defaultdict(int)
    for row in results:
        counts[row.get("Status", "ERROR")] += 1
    return {
        "total": len(results),
        "found": counts["FOUND"],
        "notFound": counts["NOT_FOUND"],
        "multipleMatch": counts["MULTIPLE_MATCH"],
        "error": counts["ERROR"],
    }


def _base_result(old_path: str, old_url: str, search_file_name: str) -> dict[str, str]:
    return {
        "OldPath": old_path,
        "OldURL": old_url,
        "SearchFileName": search_file_name,
        "MatchedFileName": "",
        "NewURL": "",
        "Status": "",
        "Remark": "",
    }


def _status_from_matches(search_file_name: str, matches: list[dict[str, str]]) -> dict[str, str]:
    if len(matches) == 1:
        match = matches[0]
        library = match.get("driveName", "")
        library_note = f" in {library}" if library else ""
        return {
            "MatchedFileName": match.get("name", ""),
            "NewURL": match.get("webUrl", ""),
            "Status": "FOUND",
            "Remark": f"Exact filename match{library_note}",
        }

    if len(matches) == 0:
        return {"Status": "NOT_FOUND", "Remark": "No exact filename match"}

    urls = [item.get("webUrl", "") for item in matches if item.get("webUrl")]
    libraries = sorted({item.get("driveName", "") for item in matches if item.get("driveName")})
    location_note = f" in {', '.join(libraries)}" if libraries else ""
    return {
        "MatchedFileName": search_file_name,
        "NewURL": " | ".join(urls),
        "Status": "MULTIPLE_MATCH",
        "Remark": f"{len(matches)} exact filename matches found{location_note}",
    }
