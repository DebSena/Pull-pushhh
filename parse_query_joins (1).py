"""
parse_query_joins.py
────────────────────────────────────────────────────────────────────────────────
Parses database query logs from a CSV file, extracts table-join relationships
from SQL text using sqlglot AST traversal, and outputs an aggregated
join-pair breakdown to a plain Excel workbook.

Usage
-----
    python parse_query_joins.py --input query_logs.csv --output join_analysis.xlsx

Dependencies
------------
    pip install sqlglot pandas openpyxl
"""

import argparse
import logging
import re
import sys
from typing import Optional

import pandas as pd
import sqlglot
import sqlglot.expressions as exp

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

SERVICE_PREFIXES = ("SVP", "SVR", "SVC", "SVT", "ETL", "DBA", "OP")

_BLOCK_COMMENT = re.compile(r"/\*.*?\*/", re.DOTALL)
_LINE_COMMENT   = re.compile(r"(--[^\n]*|//[^\n]*)", re.MULTILINE)


def strip_comments(sql: str) -> str:
    sql = _BLOCK_COMMENT.sub(" ", sql)
    sql = _LINE_COMMENT.sub("", sql)
    return sql.strip()


def clean_table_name(name: str) -> str:
    parts = name.split(".")
    return parts[-1].strip("`\"' ")


def extract_join_pairs(sql_text: str) -> list[dict]:
    pairs: list[dict] = []
    cleaned = strip_comments(sql_text)
    if not cleaned:
        return pairs

    try:
        statements = sqlglot.parse(cleaned)
    except Exception as exc:
        raise ValueError(f"sqlglot.parse failed: {exc}") from exc

    for statement in statements:
        if statement is None:
            continue

        alias_map: dict[str, str] = {}
        for tbl in statement.find_all(exp.Table):
            physical = clean_table_name(tbl.name or "")
            alias = tbl.alias or ""
            if alias:
                alias_map[alias.strip()] = physical

        def resolve(name: str) -> str:
            name = name.strip()
            return alias_map.get(name, clean_table_name(name))

        for select in statement.find_all(exp.Select):
            from_clause = select.find(exp.From)
            if from_clause is None:
                continue
            from_tbl = from_clause.find(exp.Table)
            if from_tbl is None:
                continue

            left_anchor = resolve(from_tbl.alias or from_tbl.name or "")
            if not left_anchor:
                continue

            for join_node in select.find_all(exp.Join):
                joined_tbl = join_node.find(exp.Table)
                if joined_tbl is None:
                    continue
                right = resolve(joined_tbl.alias or joined_tbl.name or "")
                if not right:
                    continue
                if left_anchor and right and left_anchor != right:
                    pairs.append({"left": left_anchor, "right": right})
                left_anchor = right

    return pairs


def get_app_prefix(user_name: str) -> Optional[str]:
    u = (user_name or "").upper()
    for pfx in SERVICE_PREFIXES:
        if u.startswith(pfx):
            return pfx
    return None


def build_join_dataframe(input_csv: str) -> pd.DataFrame:
    log.info("Reading input CSV: %s", input_csv)
    df_raw = pd.read_csv(input_csv, parse_dates=["LogDate", "StartTime", "LastResponseTime"])

    required_cols = {"user_name", "Db_nm", "Tbl_nm", "SqlTextInfo",
                     "LogDate", "StartTime", "LastResponseTime"}
    missing = required_cols - set(df_raw.columns)
    if missing:
        raise ValueError(f"Input CSV is missing required columns: {missing}")

    log.info("Loaded %d query-log rows.", len(df_raw))

    df_raw["date_wid"]    = df_raw["LogDate"].dt.strftime("%Y%m%d").astype(int)
    df_raw["Metric date"] = df_raw["LogDate"].dt.date
    df_raw["runtime_sec"] = (
        df_raw["LastResponseTime"] - df_raw["StartTime"]
    ).dt.total_seconds()
    df_raw["app_prefix"]  = df_raw["user_name"].apply(get_app_prefix)

    exploded_rows: list[dict] = []
    skipped = 0

    for idx, row in df_raw.iterrows():
        sql_text = str(row.get("SqlTextInfo", "") or "")
        try:
            pairs = extract_join_pairs(sql_text)
        except Exception as exc:
            log.warning("Row %d: skipping unparseable SQL — %s", idx, exc)
            skipped += 1
            continue

        if not pairs:
            continue

        for pair in pairs:
            exploded_rows.append({
                "date_wid":              row["date_wid"],
                "Metric date":           row["Metric date"],
                "left_join_table_name":  pair["left"],
                "right_join_table_name": pair["right"],
                "user_name":             row["user_name"],
                "app_prefix":            row["app_prefix"],
                "runtime_sec":           row["runtime_sec"],
            })

    if skipped:
        log.warning("Skipped %d rows due to SQL parse errors.", skipped)

    if not exploded_rows:
        raise RuntimeError("No join pairs could be extracted from the input data.")

    df_exp = pd.DataFrame(exploded_rows)
    log.info("Exploded into %d join-pair rows before aggregation.", len(df_exp))

    group_keys = ["date_wid", "Metric date",
                  "left_join_table_name", "right_join_table_name"]

    agg = (
        df_exp.groupby(group_keys, as_index=False)
        .agg(
            join_count   = ("left_join_table_name", "count"),
            unique_users = ("user_name",   pd.Series.nunique),
            unique_app   = ("app_prefix",  pd.Series.nunique),
            query_count  = ("user_name",   "count"),
            avg_runtime  = ("runtime_sec", "mean"),
        )
    )

    agg["avg_runtime"] = agg["avg_runtime"].round(2)
    agg.sort_values(by=["date_wid", "join_count"], ascending=[True, False], inplace=True)
    agg.reset_index(drop=True, inplace=True)
    agg.insert(0, "Row_id", range(1, len(agg) + 1))

    final_cols = [
        "Row_id", "date_wid", "Metric date",
        "left_join_table_name", "right_join_table_name",
        "join_count", "unique_users", "unique_app",
        "query_count", "avg_runtime",
    ]
    agg = agg[final_cols]

    log.info("Final aggregated output: %d rows × %d columns.", *agg.shape)
    return agg


def write_excel(df: pd.DataFrame, output_path: str) -> None:
    log.info("Writing Excel output: %s", output_path)
    with pd.ExcelWriter(output_path, engine="openpyxl", date_format="YYYY-MM-DD") as writer:
        df.to_excel(writer, sheet_name="Join Analysis", index=False)
    log.info("Excel workbook saved successfully → %s", output_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Parse SQL query logs and output a join-pair analysis Excel file."
    )
    parser.add_argument("--input",  "-i", default="query_logs.csv",
                        help="Path to the input CSV file (default: query_logs.csv)")
    parser.add_argument("--output", "-o", default="join_analysis.xlsx",
                        help="Path for the output Excel file (default: join_analysis.xlsx)")
    args = parser.parse_args()

    df_result = build_join_dataframe(args.input)
    write_excel(df_result, args.output)

    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 140)
    print("\n── Preview (first 10 rows) ──────────────────────────────────")
    print(df_result.head(10).to_string(index=False))
    print(f"\n✓ Full output → {args.output}")


if __name__ == "__main__":
    main()
