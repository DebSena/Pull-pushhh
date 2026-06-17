"""
Teradata SQL Log Converter (Slim)
==================================
Reads an Excel / CSV file of Teradata SQL logs and converts each query into
a 7-column data-product lineage schema.

Input columns  : Usr_nm, obj_db_nm, obj_tbl_nm, SqlTextInfo, LogDate
Output columns : row_wid (auto-generated UUID), date_wid, metric_date,
                 left_data_product_name, right_data_product_name,
                 join_count, unique_users
"""

import re
import uuid
import warnings
import logging
import argparse
from pathlib import Path
from typing import Optional

import pandas as pd
import sqlglot
from sqlglot import exp

warnings.filterwarnings("ignore")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def make_date_wid(date_val) -> str:
    """YYYYMMDD string used as date_wid."""
    try:
        dt = pd.to_datetime(date_val)
        return dt.strftime("%Y%m%d")
    except Exception:
        return "00000000"


def make_metric_date(date_val) -> str:
    """Human-readable YYYY-MM-DD used as metric_date."""
    try:
        dt = pd.to_datetime(date_val)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return ""


def normalize_name(name: Optional[str]) -> str:
    if not name:
        return ""
    return name.strip().strip('"').strip("'").lower()


def qualify(db: Optional[str], tbl: Optional[str]) -> str:
    """Return db.table or just table."""
    db = normalize_name(db)
    tbl = normalize_name(tbl)
    if db and tbl:
        return f"{db}.{tbl}"
    return tbl or db or ""


# ---------------------------------------------------------------------------
# SQL Cleaner  –  make Teradata SQL palatable for sqlglot
# ---------------------------------------------------------------------------

TERADATA_CLEANUP = [
    (r"\bFORMAT\s+'[^']*'", ""),
    (r"\bTITLE\s+'[^']*'", ""),
    (r"\bNAMED\s+\w+", ""),
    (r"\bCASEWORD\s+\w+", ""),
    (r"\bAT\s+ISOLATION\s+LEVEL\s+\w+", ""),
    (r"\bLOCKING\s+(?:ROW|TABLE|DATABASE|VIEW)\s+(?:FOR\s+)?(?:ACCESS|WRITE|READ|EXCLUSIVE)", ""),
    (r"\bLOCKING\s+\S+\s+FOR\s+(?:ACCESS|WRITE|READ|EXCLUSIVE)", ""),
    (r"\b(?:COMPRESS|NO\s+COMPRESS)\b[^,;)]*", ""),
    (r"\bNORMALIZE(?:\s+ON\s+MEETS\s+OR\s+OVERLAPS)?", ""),
    (r"\bEXPAND\s+ON\s+\w+", ""),
    (r"--[^\n]*", ""),
    (r"/\*.*?\*/", " ", re.DOTALL),
    (r"\bBTEQ\b.*", ""),
    (r";\s*$", ""),
    (r"\s+", " "),
]


def clean_sql(sql: str) -> str:
    s = str(sql).strip()
    for item in TERADATA_CLEANUP:
        flags = item[2] if len(item) == 3 else 0
        s = re.sub(item[0], item[1], s, flags=flags | re.IGNORECASE)
    return s.strip()


# ---------------------------------------------------------------------------
# Fallback regex-based join extractor
# ---------------------------------------------------------------------------

JOIN_RE = re.compile(
    r"""
    (?:FROM|JOIN)\s+
    (?:([\w$]+)\s*\.\s*)?([\w$]+)
    (?:\s+AS\s+([\w$]+)|\s+([\w$]+))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

JOIN_CLAUSE_RE = re.compile(
    r"""
    ((?:LEFT|RIGHT|FULL|CROSS|INNER|OUTER)\s+(?:OUTER\s+)?)?
    JOIN\s+
    (?:([\w$]+)\s*\.\s*)?([\w$]+)
    (?:\s+AS\s+([\w$]+)|\s+([\w$]+))?
    \s+ON\s+(.+?)
    (?=(?:LEFT|RIGHT|FULL|CROSS|INNER|OUTER)?\s*JOIN|\s+WHERE|\s+GROUP|\s+ORDER|\s+HAVING|$)
    """,
    re.IGNORECASE | re.VERBOSE | re.DOTALL,
)

ON_KEY_RE = re.compile(r"([\w$.]+)\s*=\s*([\w$.]+)", re.IGNORECASE)


def extract_table_name(alias_map: dict, ref: str) -> str:
    ref = normalize_name(ref)
    return alias_map.get(ref, ref)


def regex_extract_joins(sql: str, fallback_db: str = "") -> list[dict]:
    alias_map = {}
    tables_in_order = []
    for m in JOIN_RE.finditer(sql):
        db = m.group(1) or fallback_db
        tbl = m.group(2)
        alias = m.group(3) or m.group(4) or tbl
        qualified = qualify(db, tbl)
        alias_map[normalize_name(alias)] = qualified
        alias_map[normalize_name(tbl)] = qualified
        tables_in_order.append(qualified)

    joins = []
    for m in JOIN_CLAUSE_RE.finditer(sql):
        db = m.group(2) or fallback_db
        tbl = m.group(3)
        alias = m.group(4) or m.group(5) or tbl
        on_clause = m.group(6)
        right_table = qualify(db, tbl)
        alias_map[normalize_name(alias)] = right_table

        keys = ON_KEY_RE.findall(on_clause)
        left_table = None
        join_keys = []
        for lk, rk in keys:
            l_parts = lk.split(".")
            r_parts = rk.split(".")
            l_tbl = extract_table_name(alias_map, l_parts[0]) if len(l_parts) > 1 else None
            r_tbl = extract_table_name(alias_map, r_parts[0]) if len(r_parts) > 1 else None
            l_col = l_parts[-1]
            r_col = r_parts[-1]
            if r_tbl and r_tbl == right_table:
                left_table = l_tbl
                join_keys.append({"left_key": l_col, "right_key": r_col})
            elif l_tbl and l_tbl == right_table:
                left_table = r_tbl
                join_keys.append({"left_key": r_col, "right_key": l_col})
            else:
                join_keys.append({"left_key": l_col, "right_key": r_col})

        if not left_table and len(tables_in_order) >= 2:
            idx = tables_in_order.index(right_table) if right_table in tables_in_order else -1
            left_table = tables_in_order[idx - 1] if idx > 0 else tables_in_order[0]

        joins.append({
            "left_table":  left_table or "",
            "right_table": right_table,
            "join_keys":   join_keys,
        })

    return joins


# ---------------------------------------------------------------------------
# AST-based join extractor via sqlglot
# ---------------------------------------------------------------------------

def ast_extract_joins(sql: str, fallback_db: str = "") -> list[dict]:
    try:
        statements = sqlglot.parse(sql, dialect="teradata", error_level=sqlglot.ErrorLevel.IGNORE)
    except Exception:
        statements = []

    if not statements:
        return []

    joins_out = []
    for stmt in statements:
        if stmt is None:
            continue
        alias_map = {}
        for tbl_expr in stmt.find_all(exp.Table):
            db   = tbl_expr.args.get("db")
            name = tbl_expr.args.get("this")
            alias = tbl_expr.args.get("alias")
            db_s  = db.name if db else fallback_db
            tbl_s = name.name if name else ""
            alias_s = alias.name if alias else tbl_s
            qualified = qualify(db_s, tbl_s)
            alias_map[normalize_name(alias_s)] = qualified
            alias_map[normalize_name(tbl_s)]   = qualified

        def resolve(node) -> str:
            if isinstance(node, exp.Table):
                db   = node.args.get("db")
                name = node.args.get("this")
                db_s  = db.name if db else fallback_db
                tbl_s = name.name if name else ""
                return qualify(db_s, tbl_s)
            if isinstance(node, exp.Column):
                tbl = node.args.get("table")
                if tbl:
                    tname = normalize_name(tbl.name)
                    return alias_map.get(tname, tname)
            return ""

        for select in stmt.find_all(exp.Select):
            from_clause = select.args.get("from")
            if not from_clause:
                continue
            from_tbl_expr = from_clause.find(exp.Table)
            if from_tbl_expr is None:
                continue
            left_anchor = resolve(from_tbl_expr)

            for join_node in select.args.get("joins", []):
                right_expr  = join_node.args.get("this")
                right_table = resolve(right_expr) if right_expr else ""

                on_node   = join_node.args.get("on")
                join_keys = []
                if on_node:
                    for eq_node in on_node.find_all(exp.EQ):
                        lc = eq_node.args.get("this")
                        rc = eq_node.args.get("expression")
                        if lc and rc:
                            join_keys.append({
                                "left_key":  lc.name if hasattr(lc, "name") else str(lc),
                                "right_key": rc.name if hasattr(rc, "name") else str(rc),
                            })

                joins_out.append({
                    "left_table":  left_anchor,
                    "right_table": right_table,
                    "join_keys":   join_keys,
                })
                left_anchor = right_table or left_anchor

    return joins_out


# ---------------------------------------------------------------------------
# Master join extractor
# ---------------------------------------------------------------------------

def extract_joins(sql: str, fallback_db: str = "") -> list[dict]:
    """Try AST first; fall back to regex for unparseable queries."""
    cleaned = clean_sql(sql)
    joins = []
    try:
        joins = ast_extract_joins(cleaned, fallback_db)
    except Exception as e:
        log.debug(f"AST extraction failed ({e}), using regex fallback")

    if not joins:
        joins = regex_extract_joins(cleaned, fallback_db)

    # Deduplicate while preserving order; keep only pairs with both sides
    seen   = set()
    unique = []
    for j in joins:
        key = (j["left_table"], j["right_table"])
        if key not in seen and j["left_table"] and j["right_table"]:
            seen.add(key)
            unique.append(j)
    return unique


# ---------------------------------------------------------------------------
# Core transformation  –  7-column output
# ---------------------------------------------------------------------------

def transform_logs(df: pd.DataFrame) -> pd.DataFrame:
    """
    Accepts the raw log DataFrame and returns a slim 7-column DataFrame:
        row_wid, date_wid, metric_date,
        left_data_product_name, right_data_product_name,
        join_count, unique_users
    """
    if "SqlTextInfo" not in df.columns and "sqltextinfo" not in [c.lower() for c in df.columns]:
        raise ValueError("Missing required column: SqlTextInfo")

    col_map = {c.lower(): c for c in df.columns}

    def get_col(name: str, default=""):
        mapped = col_map.get(name.lower())
        return df[mapped] if mapped else pd.Series([default] * len(df))

    usr_nm    = get_col("usr_nm")
    obj_db_nm = get_col("obj_db_nm")
    obj_tbl_nm = get_col("obj_tbl_nm")
    sql_col   = get_col("sqltextinfo")
    log_date  = get_col("logdate")

    # ------------------------------------------------------------------
    # Pass 1: extract per-query join pairs
    # ------------------------------------------------------------------
    records = []

    for idx in range(len(df)):
        sql  = str(sql_col.iloc[idx])
        usr  = str(usr_nm.iloc[idx])
        db   = str(obj_db_nm.iloc[idx])
        tbl  = str(obj_tbl_nm.iloc[idx])
        date = log_date.iloc[idx]

        if not sql or sql.lower() in ("nan", "none", ""):
            continue

        joins = extract_joins(sql, fallback_db=db)

        if not joins:
            # No JOIN found – record single-table hit with empty right side
            qualified = qualify(db, tbl)
            if qualified:
                records.append({
                    "left_table":  qualified,
                    "right_table": "",
                    "join_keys":   [],
                    "usr":         usr,
                    "date":        date,
                })
            continue

        for j in joins:
            records.append({
                "left_table":  j["left_table"] or qualify(db, tbl),
                "right_table": j["right_table"],
                "join_keys":   j["join_keys"],
                "usr":         usr,
                "date":        date,
            })

    if not records:
        log.warning("No records extracted. Output will be empty.")
        return pd.DataFrame()

    raw_df = pd.DataFrame(records)

    # ------------------------------------------------------------------
    # Pass 2: aggregate per (left_table, right_table, date)
    # ------------------------------------------------------------------
    raw_df["date_key"] = raw_df["date"].apply(make_date_wid)

    agg_rows = []
    for (left, right, date_key), grp in raw_df.groupby(
        ["left_table", "right_table", "date_key"], sort=False
    ):
        # join_count = total distinct ON-key pairs across the group
        all_keys   = [k for jk in grp["join_keys"] for k in jk]
        join_count = max(len(all_keys), 1)

        unique_users = grp["usr"].nunique()

        # metric_date: recover a human-readable date from the first row in group
        sample_date = grp["date"].iloc[0]

        agg_rows.append({
            "row_wid":                  str(uuid.uuid4()).replace("-", "")[:16].upper(),
            "date_wid":                 date_key,
            "metric_date":              make_metric_date(sample_date),
            "left_data_product_name":   left,
            "right_data_product_name":  right if right else "",
            "join_count":               join_count,
            "unique_users":             unique_users,
        })

    return pd.DataFrame(agg_rows)


# ---------------------------------------------------------------------------
# I/O
# ---------------------------------------------------------------------------

def read_input(path: str, sheet=0) -> pd.DataFrame:
    p   = Path(path)
    ext = p.suffix.lower()
    log.info(f"Reading input: {path}")
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    elif ext == ".csv":
        df = pd.read_csv(path, dtype=str)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
    log.info(f"Loaded {len(df)} rows, columns: {list(df.columns)}")
    return df


def write_output(df: pd.DataFrame, out_path: str):
    p   = Path(out_path)
    ext = p.suffix.lower()
    if ext in (".xlsx", ".xls"):
        df.to_excel(out_path, index=False)
    else:
        df.to_csv(out_path, index=False)
    log.info(f"Output written to {out_path}  ({len(df)} rows)")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Convert Teradata SQL logs → slim 7-column lineage schema"
    )
    parser.add_argument("input",  help="Path to input Excel/CSV file")
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Output path (default: <input>_slim.xlsx)",
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="Sheet name or index for Excel input (default: 0)",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug logging",
    )
    args = parser.parse_args()

    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)

    if args.output is None:
        p = Path(args.input)
        args.output = str(p.parent / f"{p.stem}_slim.xlsx")

    df_in  = read_input(args.input, sheet=args.sheet)
    df_out = transform_logs(df_in)

    if df_out.empty:
        log.warning("No data produced. Check your input file.")
        return

    write_output(df_out, args.output)
    print(f"\n✅  Done!  {len(df_out)} lineage rows written to: {args.output}")
    print("\nOutput columns:", list(df_out.columns))
    print("\nSample output:")
    print(df_out.head(5).to_string(index=False))


if __name__ == "__main__":
    main()
