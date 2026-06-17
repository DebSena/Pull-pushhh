"""
Teradata SQL Log Converter
==========================
Reads an Excel file of Teradata SQL logs and converts each query into
a structured data-product lineage schema, capturing all JOIN relationships
(including multi-key joins), user metrics, and scoring.

Input schema  : Usr_nm, obj_db_nm, obj_tbl_nm, SqlTextInfo, LogDate
Output schema : row_wid, date_wid, metric_wid, left_data_product_name,
                left_data_product_id, right_data_product_name,
                right_data_product_id, join_count, unique_users, unique_app,
                query_count, avg_runtime, recommendation_score,
                domain_similarity_score, created_timestamp
"""

import re
import hashlib
import uuid
import warnings
import logging
import argparse
from collections import defaultdict
from datetime import datetime, timezone
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

def make_wid(seed: str) -> str:
    """Deterministic 12-char hex WID from a seed string."""
    return hashlib.md5(seed.encode()).hexdigest()[:12].upper()


def make_date_wid(date_val) -> str:
    """YYYYMMDD integer-style WID."""
    try:
        dt = pd.to_datetime(date_val)
        return dt.strftime("%Y%m%d")
    except Exception:
        return "00000000"


def make_metric_wid(left: str, right: str, date_str: str) -> str:
    return make_wid(f"{left}|{right}|{date_str}")


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
    # Remove TD-specific keywords that confuse the parser
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
    (r"--[^\n]*", ""),           # single-line comments
    (r"/\*.*?\*/", " ", re.DOTALL),  # block comments
    (r"\bBTEQ\b.*", ""),
    (r";\s*$", ""),              # trailing semicolons
    (r"\s+", " "),               # collapse whitespace
]


def clean_sql(sql: str) -> str:
    s = str(sql).strip()
    for item in TERADATA_CLEANUP:
        flags = item[2] if len(item) == 3 else 0
        s = re.sub(item[0], item[1], s, flags=flags | re.IGNORECASE)
    return s.strip()


# ---------------------------------------------------------------------------
# Fallback regex-based join extractor (when sqlglot fails)
# ---------------------------------------------------------------------------

JOIN_RE = re.compile(
    r"""
    (?:FROM|JOIN)\s+
    (?:([\w$]+)\s*\.\s*)?([\w$]+)   # [db.]table
    (?:\s+AS\s+([\w$]+)|\s+([\w$]+))? # optional alias
    """,
    re.IGNORECASE | re.VERBOSE,
)

JOIN_CLAUSE_RE = re.compile(
    r"""
    ((?:LEFT|RIGHT|FULL|CROSS|INNER|OUTER)\s+(?:OUTER\s+)?)?
    JOIN\s+
    (?:([\w$]+)\s*\.\s*)?([\w$]+)   # [db.]table
    (?:\s+AS\s+([\w$]+)|\s+([\w$]+))? # optional alias
    \s+ON\s+(.+?)                    # ON clause
    (?=(?:LEFT|RIGHT|FULL|CROSS|INNER|OUTER)?\s*JOIN|\s+WHERE|\s+GROUP|\s+ORDER|\s+HAVING|$)
    """,
    re.IGNORECASE | re.VERBOSE | re.DOTALL,
)

ON_KEY_RE = re.compile(
    r"([\w$.]+)\s*=\s*([\w$.]+)",
    re.IGNORECASE,
)


def extract_table_name(alias_map: dict, ref: str) -> str:
    """Resolve an alias or qualified name to its canonical table string."""
    ref = normalize_name(ref)
    return alias_map.get(ref, ref)


def regex_extract_joins(sql: str, fallback_db: str = "") -> list[dict]:
    """
    Rough but resilient join extractor using regex when sqlglot fails.
    Returns list of join pair dicts.
    """
    # Build alias→table map from all FROM / JOIN clauses
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
    # Parse each JOIN ... ON clause
    for m in JOIN_CLAUSE_RE.finditer(sql):
        db = m.group(2) or fallback_db
        tbl = m.group(3)
        alias = m.group(4) or m.group(5) or tbl
        on_clause = m.group(6)
        right_table = qualify(db, tbl)
        alias_map[normalize_name(alias)] = right_table

        # Find the "left" table: the one referenced in the ON keys that is NOT the right table
        keys = ON_KEY_RE.findall(on_clause)
        left_table = None
        join_keys = []
        for lk, rk in keys:
            # each side may be alias.column
            l_parts = lk.split(".")
            r_parts = rk.split(".")
            l_tbl = extract_table_name(alias_map, l_parts[0]) if len(l_parts) > 1 else None
            r_tbl = extract_table_name(alias_map, r_parts[0]) if len(r_parts) > 1 else None

            l_col = l_parts[-1]
            r_col = r_parts[-1]

            # Determine which side is left vs right
            if r_tbl and r_tbl == right_table:
                left_table = l_tbl
                join_keys.append({"left_key": l_col, "right_key": r_col})
            elif l_tbl and l_tbl == right_table:
                left_table = r_tbl
                join_keys.append({"left_key": r_col, "right_key": l_col})
            else:
                join_keys.append({"left_key": l_col, "right_key": r_col})

        if not left_table and len(tables_in_order) >= 2:
            # Fallback: use the table that appeared just before this join
            idx = tables_in_order.index(right_table) if right_table in tables_in_order else -1
            left_table = tables_in_order[idx - 1] if idx > 0 else tables_in_order[0]

        joins.append({
            "left_table": left_table or "",
            "right_table": right_table,
            "join_keys": join_keys,
            "join_type": (m.group(1) or "INNER").strip().upper(),
        })

    return joins


# ---------------------------------------------------------------------------
# AST-based join extractor via sqlglot
# ---------------------------------------------------------------------------

def ast_extract_joins(sql: str, fallback_db: str = "") -> list[dict]:
    """
    Use sqlglot to parse and walk the AST, collecting all Join nodes.
    Returns the same structure as regex_extract_joins.
    """
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
        # Build alias map for this statement
        alias_map = {}
        for tbl_expr in stmt.find_all(exp.Table):
            db = tbl_expr.args.get("db")
            name = tbl_expr.args.get("this")
            alias = tbl_expr.args.get("alias")
            db_s = db.name if db else fallback_db
            tbl_s = name.name if name else ""
            alias_s = alias.name if alias else tbl_s
            qualified = qualify(db_s, tbl_s)
            alias_map[normalize_name(alias_s)] = qualified
            alias_map[normalize_name(tbl_s)] = qualified

        def resolve(node) -> str:
            if isinstance(node, exp.Table):
                db = node.args.get("db")
                name = node.args.get("this")
                db_s = db.name if db else fallback_db
                tbl_s = name.name if name else ""
                return qualify(db_s, tbl_s)
            if isinstance(node, exp.Column):
                tbl = node.args.get("table")
                if tbl:
                    tname = normalize_name(tbl.name)
                    return alias_map.get(tname, tname)
            return ""

        # Walk all Select nodes (handles CTEs, subqueries)
        for select in stmt.find_all(exp.Select):
            from_clause = select.args.get("from")
            if not from_clause:
                continue
            # The "anchor" (left) table is the FROM table
            from_tbl_expr = from_clause.find(exp.Table)
            if from_tbl_expr is None:
                continue
            left_anchor = resolve(from_tbl_expr)

            for join_node in select.args.get("joins", []):
                join_kind = join_node.args.get("kind", "")
                join_side = join_node.args.get("side", "")
                jtype = f"{join_side} {join_kind}".strip() or "INNER"

                right_expr = join_node.args.get("this")
                right_table = resolve(right_expr) if right_expr else ""

                # ON condition
                on_node = join_node.args.get("on")
                join_keys = []
                if on_node:
                    for eq_node in on_node.find_all(exp.EQ):
                        left_col = eq_node.args.get("this")
                        right_col = eq_node.args.get("expression")
                        if left_col and right_col:
                            lc = left_col.name if hasattr(left_col, "name") else str(left_col)
                            rc = right_col.name if hasattr(right_col, "name") else str(right_col)
                            join_keys.append({"left_key": lc, "right_key": rc})

                joins_out.append({
                    "left_table": left_anchor,
                    "right_table": right_table,
                    "join_keys": join_keys,
                    "join_type": jtype.upper(),
                })
                # Subsequent joins in the same SELECT chain against prior joins
                left_anchor = right_table or left_anchor

    return joins_out


# ---------------------------------------------------------------------------
# Master join extractor (tries AST, falls back to regex)
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

    # Deduplicate while preserving order
    seen = set()
    unique = []
    for j in joins:
        key = (j["left_table"], j["right_table"])
        if key not in seen and j["left_table"] and j["right_table"]:
            seen.add(key)
            unique.append(j)
    return unique


# ---------------------------------------------------------------------------
# Scoring helpers
# ---------------------------------------------------------------------------

DOMAIN_KEYWORDS = {
    "finance":  ["fin", "account", "ledger", "payment", "invoice", "revenue", "cost", "budget"],
    "customer": ["cust", "client", "member", "user", "subscriber", "contact"],
    "product":  ["prod", "item", "sku", "catalog", "inventory", "stock"],
    "order":    ["order", "sale", "transaction", "purchase", "cart"],
    "hr":       ["emp", "staff", "payroll", "department", "position"],
    "logistics":["ship", "deliver", "warehouse", "carrier", "freight", "dispatch"],
    "marketing":["campaign", "promo", "lead", "segment", "channel"],
    "analytics":["dim", "fact", "agg", "metric", "report", "kpi", "mart"],
}


def infer_domain(table_name: str) -> str:
    t = table_name.lower()
    for domain, kws in DOMAIN_KEYWORDS.items():
        if any(kw in t for kw in kws):
            return domain
    return "unknown"


def domain_similarity(left: str, right: str) -> float:
    """0.0–1.0: how similar the inferred domains are."""
    dl = infer_domain(left)
    dr = infer_domain(right)
    if dl == "unknown" or dr == "unknown":
        return 0.3   # neutral
    return 1.0 if dl == dr else 0.2


def recommendation_score(
    join_count: int,
    query_count: int,
    unique_users: int,
    domain_sim: float,
) -> float:
    """
    Composite score 0–100 reflecting how 'interesting' this lineage pair is.
    Higher = more frequently joined, more users, same domain.
    """
    freq_score   = min(query_count / 10.0, 10) * 4     # up to 40
    join_score   = min(join_count / 5.0, 5)  * 4       # up to 20
    user_score   = min(unique_users / 3.0, 10) * 2     # up to 20
    domain_score = domain_sim * 20                      # up to 20
    return round(freq_score + join_score + user_score + domain_score, 2)


# ---------------------------------------------------------------------------
# Core transformation
# ---------------------------------------------------------------------------

def transform_logs(df: pd.DataFrame) -> pd.DataFrame:
    """
    Main transformation function.
    Accepts the raw log DataFrame and returns the output schema DataFrame.
    """
    required = {"SqlTextInfo"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Normalize column names (case-insensitive lookup)
    col_map = {c.lower(): c for c in df.columns}

    def get_col(name: str, default=""):
        mapped = col_map.get(name.lower())
        return df[mapped] if mapped else pd.Series([default] * len(df))

    usr_nm      = get_col("usr_nm")
    obj_db_nm   = get_col("obj_db_nm")
    obj_tbl_nm  = get_col("obj_tbl_nm")
    sql_col     = get_col("sqltextinfo")
    log_date    = get_col("logdate")

    # -----------------------------------------------------------------------
    # Pass 1: extract per-query join pairs
    # -----------------------------------------------------------------------
    records = []   # one record per (left_table, right_table, date, user)

    for idx, row in df.iterrows():
        sql   = str(sql_col.iloc[idx]) if isinstance(sql_col, pd.Series) else str(row.get("SqlTextInfo", ""))
        usr   = str(usr_nm.iloc[idx])  if isinstance(usr_nm, pd.Series)  else ""
        db    = str(obj_db_nm.iloc[idx]) if isinstance(obj_db_nm, pd.Series) else ""
        tbl   = str(obj_tbl_nm.iloc[idx]) if isinstance(obj_tbl_nm, pd.Series) else ""
        date  = log_date.iloc[idx] if isinstance(log_date, pd.Series) else None

        if not sql or sql.lower() in ("nan", "none", ""):
            continue

        joins = extract_joins(sql, fallback_db=db)

        if not joins:
            # Even without a JOIN we record the single table hit
            qualified = qualify(db, tbl)
            if qualified:
                records.append({
                    "left_table":  qualified,
                    "right_table": "",
                    "join_keys":   [],
                    "join_type":   "NONE",
                    "usr":         usr,
                    "date":        date,
                    "sql":         sql,
                })
            continue

        for j in joins:
            records.append({
                "left_table":  j["left_table"] or qualify(db, tbl),
                "right_table": j["right_table"],
                "join_keys":   j["join_keys"],
                "join_type":   j["join_type"],
                "usr":         usr,
                "date":        date,
                "sql":         sql,
            })

    if not records:
        log.warning("No join records extracted. Output will be empty.")
        return pd.DataFrame()

    raw_df = pd.DataFrame(records)

    # -----------------------------------------------------------------------
    # Pass 2: aggregate per (left_table, right_table, date)
    # -----------------------------------------------------------------------
    raw_df["date_key"] = raw_df["date"].apply(make_date_wid)
    raw_df["pair_key"] = raw_df["left_table"] + "||" + raw_df["right_table"]

    agg_rows = []
    for (left, right, date_key), grp in raw_df.groupby(
        ["left_table", "right_table", "date_key"], sort=False
    ):
        unique_users = grp["usr"].nunique()
        unique_apps  = grp["usr"].apply(lambda u: u.split("_")[0] if "_" in u else u).nunique()
        query_count  = len(grp)
        # join_count   = distinct ON-key pairs across the group
        all_keys = [k for jk in grp["join_keys"] for k in jk]
        join_count = max(len(all_keys), 1)
        # avg_runtime  = placeholder (no runtime in source; set to -1 if absent)
        avg_runtime  = -1.0

        dom_sim   = domain_similarity(left, right)
        rec_score = recommendation_score(join_count, query_count, unique_users, dom_sim)

        left_id   = make_wid(left)
        right_id  = make_wid(right) if right else ""
        row_wid   = str(uuid.uuid4()).replace("-", "")[:16].upper()
        metric_wid = make_metric_wid(left, right, date_key)

        agg_rows.append({
            "row_wid":                   row_wid,
            "date_wid":                  date_key,
            "metric_wid":                metric_wid,
            "left_data_product_name":    left,
            "left_data_product_id":      left_id,
            "right_data_product_name":   right if right else "",
            "right_data_product_id":     right_id,
            "join_count":                join_count,
            "unique_users":              unique_users,
            "unique_app":                unique_apps,
            "query_count":               query_count,
            "avg_runtime":               avg_runtime,
            "recommendation_score":      rec_score,
            "domain_similarity_score":   round(dom_sim, 4),
            "created_timestamp":         datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        })

    out_df = pd.DataFrame(agg_rows)
    return out_df


# ---------------------------------------------------------------------------
# I/O
# ---------------------------------------------------------------------------

def read_input(path: str) -> pd.DataFrame:
    p = Path(path)
    ext = p.suffix.lower()
    log.info(f"Reading input: {path}")
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype=str)
    elif ext == ".csv":
        df = pd.read_csv(path, dtype=str)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
    log.info(f"Loaded {len(df)} rows, columns: {list(df.columns)}")
    return df


def write_output(df: pd.DataFrame, out_path: str):
    p = Path(out_path)
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
        description="Convert Teradata SQL logs → data-product lineage schema"
    )
    parser.add_argument("input",  help="Path to input Excel/CSV file")
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Output path (default: <input>_converted.xlsx)",
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
        args.output = str(p.parent / f"{p.stem}_converted.xlsx")

    df_in  = read_input(args.input)
    df_out = transform_logs(df_in)

    if df_out.empty:
        log.warning("No data produced. Check your input file.")
        return

    write_output(df_out, args.output)
    print(f"\n✅  Done!  {len(df_out)} lineage rows written to: {args.output}")
    print("\nSample output:")
    print(df_out.head(5).to_string(index=False))


if __name__ == "__main__":
    main()
