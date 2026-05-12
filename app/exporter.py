from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.matcher import OUTPUT_COLUMNS


def export_results_csv(results: list[dict[str, str]], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(results)
    for column in OUTPUT_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    df[OUTPUT_COLUMNS].to_csv(output_path, index=False, encoding="utf-8-sig")
    return output_path
