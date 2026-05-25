from __future__ import annotations

import os
import re
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import pandas as pd

try:
    import psycopg2
    from psycopg2.extras import Json
except ImportError:  # pragma: no cover - handled at runtime with a clear error
    psycopg2 = None
    Json = None


APP_KIND = "app"
CITY_KIND = "city"
MASTER_SHEET = "Master Database"
URL_COL = "URL / App Name"
LANG_COL = "Language or Line Item"
JUNK_VALUES = {"nan", "none", "", "app/url", "app", "url", "site", "domain", URL_COL.lower()}

_APP_FALLBACK_PATH: Path | None = None
_CITY_FALLBACK_PATH: Path | None = None


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS reference_sheets (
    kind TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    source_file TEXT,
    sheet_order INTEGER NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    PRIMARY KEY (kind, sheet_name)
);

CREATE TABLE IF NOT EXISTS reference_rows (
    id BIGSERIAL PRIMARY KEY,
    kind TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    row_index INTEGER NOT NULL,
    payload JSONB NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT fk_reference_sheet
        FOREIGN KEY (kind, sheet_name)
        REFERENCES reference_sheets (kind, sheet_name)
        ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_reference_rows_lookup
    ON reference_rows (kind, sheet_name, row_index);
"""


def _normalize_text(val: Any) -> str:
    return re.sub(r"\s+", " ", str(val or "").strip().lower())


def _clean_url(u: Any) -> str:
    s = str(u or "").lower().strip()
    for prefix in ("https://", "http://", "www."):
        if s.startswith(prefix):
            s = s[len(prefix):]
    return s.rstrip("/")


def _is_blank(val: Any) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    return str(val).strip() == ""


def _safe_float(val: Any, default: float = 0.0) -> float:
    if val is None:
        return default
    try:
        if pd.isna(val):
            return default
    except Exception:
        pass
    s = str(val).strip().replace(",", "").replace(" ", "")
    if s.lower() in {"", "nan", "none", "<na>", "na", "null"}:
        return default
    if s.endswith("%"):
        try:
            return float(s[:-1]) / 100.0
        except ValueError:
            return default
    try:
        return float(s)
    except ValueError:
        return default


def _database_url() -> str:
    url = os.getenv("DATABASE_URL", "").strip()
    if url:
        return url

    host = os.getenv("PGHOST", "localhost").strip()
    port = os.getenv("PGPORT", "5432").strip()
    name = os.getenv("PGDATABASE", "").strip()
    user = os.getenv("PGUSER", "").strip()
    password = os.getenv("PGPASSWORD", "").strip()
    if not name or not user:
        raise RuntimeError(
            "PostgreSQL is not configured. Set DATABASE_URL or PGDATABASE/PGUSER/PGPASSWORD."
        )
    auth = f"{user}:{password}@" if password else f"{user}@"
    return f"postgresql://{auth}{host}:{port}/{name}"


def postgres_configured() -> bool:
    """Return True only when PostgreSQL settings are present."""
    url = os.getenv("DATABASE_URL", "").strip()
    if url:
        return True
    return bool(os.getenv("PGDATABASE", "").strip() and os.getenv("PGUSER", "").strip())


def configure_fallback_paths(app_excel_path: str | Path | None = None, city_excel_path: str | Path | None = None) -> None:
    """Register workbook paths used when PostgreSQL is not configured."""
    global _APP_FALLBACK_PATH, _CITY_FALLBACK_PATH
    _APP_FALLBACK_PATH = Path(app_excel_path) if app_excel_path else None
    _CITY_FALLBACK_PATH = Path(city_excel_path) if city_excel_path else None


def _require_driver() -> None:
    if psycopg2 is None:
        raise RuntimeError(
            "psycopg2 is not installed. Install psycopg2-binary to use PostgreSQL."
        )


@contextmanager
def _connect():
    _require_driver()
    conn = psycopg2.connect(_database_url())
    try:
        yield conn
    finally:
        conn.close()


def ensure_schema() -> None:
    if not postgres_configured() or psycopg2 is None:
        return
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(SCHEMA_SQL)
        conn.commit()


def _load_workbook_sheet_df(excel_path: str | Path | None, sheet_name: str, header: int | None = 0) -> pd.DataFrame:
    if not excel_path:
        return pd.DataFrame()
    excel_path = Path(excel_path)
    if not excel_path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(excel_path, sheet_name=sheet_name, header=header)
    except Exception:
        return pd.DataFrame()


def _workbook_sheet_names(excel_path: str | Path | None, exclude_summary: bool = False) -> list[str]:
    if not excel_path:
        return []
    excel_path = Path(excel_path)
    if not excel_path.exists():
        return []
    try:
        xls = pd.ExcelFile(excel_path)
        names = list(xls.sheet_names)
        if exclude_summary:
            names = [name for name in names if name.strip().lower() != "summary by state"]
        return names
    except Exception:
        return []


def _sheet_exists(conn, kind: str, sheet_name: str) -> bool:
    with conn.cursor() as cur:
        cur.execute(
            "SELECT 1 FROM reference_sheets WHERE kind = %s AND sheet_name = %s",
            (kind, sheet_name),
        )
        return cur.fetchone() is not None


def _max_sheet_order(conn, kind: str) -> int:
    with conn.cursor() as cur:
        cur.execute(
            "SELECT COALESCE(MAX(sheet_order), 0) FROM reference_sheets WHERE kind = %s",
            (kind,),
        )
        row = cur.fetchone()
    return int(row[0] or 0)


def _register_sheet(conn, kind: str, sheet_name: str, source_file: str | None = None, sheet_order: int | None = None) -> None:
    if sheet_order is None:
        sheet_order = _max_sheet_order(conn, kind) + 1
    with conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO reference_sheets (kind, sheet_name, source_file, sheet_order)
            VALUES (%s, %s, %s, %s)
            ON CONFLICT (kind, sheet_name)
            DO UPDATE SET source_file = EXCLUDED.source_file
            """,
            (kind, sheet_name, source_file, sheet_order),
        )


def _clear_sheet_rows(conn, kind: str, sheet_name: str) -> None:
    with conn.cursor() as cur:
        cur.execute(
            "DELETE FROM reference_rows WHERE kind = %s AND sheet_name = %s",
            (kind, sheet_name),
        )


def _rows_to_dataframe(rows: list[dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)


def _load_sheet_rows(conn, kind: str, sheet_name: str) -> list[dict[str, Any]]:
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT payload
            FROM reference_rows
            WHERE kind = %s AND sheet_name = %s
            ORDER BY row_index ASC
            """,
            (kind, sheet_name),
        )
        rows = cur.fetchall()
    return [row[0] for row in rows]


def _load_sheet_df(conn, kind: str, sheet_name: str) -> pd.DataFrame:
    rows = _load_sheet_rows(conn, kind, sheet_name)
    return _rows_to_dataframe(rows)


def _insert_sheet_rows(conn, kind: str, sheet_name: str, rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    _register_sheet(conn, kind, sheet_name)
    with conn.cursor() as cur:
        cur.execute(
            "SELECT COALESCE(MAX(row_index), 0) FROM reference_rows WHERE kind = %s AND sheet_name = %s",
            (kind, sheet_name),
        )
        current_max = int(cur.fetchone()[0] or 0)
        for offset, payload in enumerate(rows, 1):
            cur.execute(
                """
                INSERT INTO reference_rows (kind, sheet_name, row_index, payload)
                VALUES (%s, %s, %s, %s)
                """,
                (kind, sheet_name, current_max + offset, Json(payload)),
            )


def _normalize_row(row: pd.Series | dict[str, Any]) -> dict[str, Any]:
    if isinstance(row, pd.Series):
        items = row.to_dict()
    else:
        items = dict(row)
    clean: dict[str, Any] = {}
    for key, val in items.items():
        if _is_blank(val):
            clean[str(key)] = None
        elif isinstance(val, (pd.Timestamp,)):
            clean[str(key)] = val.isoformat()
        else:
            try:
                if pd.isna(val):
                    clean[str(key)] = None
                    continue
            except Exception:
                pass
            if hasattr(val, "item"):
                try:
                    clean[str(key)] = val.item()
                except Exception:
                    clean[str(key)] = val
            else:
                clean[str(key)] = val
    return clean


def import_workbook_to_postgres(excel_path: str | Path, kind: str, header: int) -> None:
    excel_path = Path(excel_path)
    if not excel_path.exists():
        return

    xls = pd.ExcelFile(excel_path)
    with _connect() as conn:
        for sheet_order, sheet_name in enumerate(xls.sheet_names, 1):
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header)
            if df.empty:
                _register_sheet(conn, kind, sheet_name, str(excel_path), sheet_order)
                continue

            if kind == APP_KIND and LANG_COL in df.columns:
                df[LANG_COL] = df[LANG_COL].ffill()

            rows: list[dict[str, Any]] = []
            for _, row in df.iterrows():
                payload = _normalize_row(row)
                if all(v is None for v in payload.values()):
                    continue
                rows.append(payload)

            _register_sheet(conn, kind, sheet_name, str(excel_path), sheet_order)
            _clear_sheet_rows(conn, kind, sheet_name)
            _insert_sheet_rows(conn, kind, sheet_name, rows)
        conn.commit()


def bootstrap_reference_data(app_excel_path: str | Path, city_excel_path: str | Path) -> None:
    configure_fallback_paths(app_excel_path, city_excel_path)
    if not postgres_configured() or psycopg2 is None:
        return
    ensure_schema()
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM reference_sheets WHERE kind = %s", (APP_KIND,))
            app_count = int(cur.fetchone()[0] or 0)
            cur.execute("SELECT COUNT(*) FROM reference_sheets WHERE kind = %s", (CITY_KIND,))
            city_count = int(cur.fetchone()[0] or 0)
        conn.commit()

    if app_count == 0:
        import_workbook_to_postgres(app_excel_path, APP_KIND, header=1)
    if city_count == 0:
        import_workbook_to_postgres(city_excel_path, CITY_KIND, header=0)


def list_sheet_names(kind: str) -> list[str]:
    if not postgres_configured() or psycopg2 is None:
        if kind == APP_KIND:
            return _workbook_sheet_names(_APP_FALLBACK_PATH)
        if kind == CITY_KIND:
            return _workbook_sheet_names(_CITY_FALLBACK_PATH, exclude_summary=True)
        return []
    with _connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT sheet_name
                FROM reference_sheets
                WHERE kind = %s
                ORDER BY sheet_order ASC, sheet_name ASC
                """,
                (kind,),
            )
            rows = cur.fetchall()
    return [row[0] for row in rows]


def load_sheet_df(kind: str, sheet_name: str) -> pd.DataFrame:
    if not postgres_configured() or psycopg2 is None:
        if kind == APP_KIND:
            return _load_workbook_sheet_df(_APP_FALLBACK_PATH, sheet_name, header=1)
        if kind == CITY_KIND:
            return _load_workbook_sheet_df(_CITY_FALLBACK_PATH, sheet_name, header=0)
        return pd.DataFrame()
    with _connect() as conn:
        return _load_sheet_df(conn, kind, sheet_name)


def load_app_sheet(sheet_name: str) -> pd.DataFrame:
    return load_sheet_df(APP_KIND, sheet_name)


def get_app_sheet_names() -> list[str]:
    return list_sheet_names(APP_KIND)


def fallback_app_sheet_names_from_workbook(excel_path: str | Path) -> list[str]:
    return _workbook_sheet_names(excel_path)


def get_app_urls_from_sheets(sheet_names_str: str) -> list[str]:
    if not sheet_names_str:
        return []
    sheet_names = [s.strip() for s in sheet_names_str.split(",") if s.strip()]
    all_urls: set[str] = set()
    for sheet_name in sheet_names:
        df = load_app_sheet(sheet_name)
        if df.empty or URL_COL not in df.columns:
            continue
        for url in df[URL_COL].dropna().astype(str).tolist():
            cleaned = _clean_url(url)
            if cleaned and cleaned not in JUNK_VALUES:
                all_urls.add(cleaned)
    return list(all_urls)


def _current_master_language(df: pd.DataFrame) -> str | None:
    if df.empty or LANG_COL not in df.columns:
        return None
    series = df[LANG_COL].dropna().astype(str)
    for value in reversed(series.tolist()):
        value = value.strip()
        if value:
            return value
    return None


def append_app_urls_to_sheet(sheet_name: str, new_urls: list[str]) -> None:
    if not new_urls:
        return

    if not postgres_configured() or psycopg2 is None:
        return

    with _connect() as conn:
        if not _sheet_exists(conn, APP_KIND, sheet_name):
            return
        df = _load_sheet_df(conn, APP_KIND, sheet_name)
        if df.empty:
            df = pd.DataFrame(columns=[URL_COL])

        columns = list(df.columns)
        if URL_COL not in columns:
            columns.append(URL_COL)

        existing_urls: set[str] = set()
        if URL_COL in df.columns:
            for value in df[URL_COL].dropna().astype(str).tolist():
                cleaned = _clean_url(value)
                if cleaned and cleaned not in JUNK_VALUES:
                    existing_urls.add(cleaned)

        current_language = _current_master_language(df) if sheet_name == MASTER_SHEET else None
        next_id = 1
        if "ID" in df.columns:
            id_values = pd.to_numeric(df["ID"], errors="coerce").dropna()
            if not id_values.empty:
                next_id = int(id_values.max()) + 1

        rows: list[dict[str, Any]] = []
        for url in new_urls:
            cleaned = _clean_url(url)
            if cleaned in JUNK_VALUES or cleaned in existing_urls:
                continue
            existing_urls.add(cleaned)
            payload = {col: None for col in columns}
            if "ID" in payload:
                payload["ID"] = next_id
                next_id += 1
            payload[URL_COL] = url
            if LANG_COL in payload and current_language:
                payload[LANG_COL] = current_language
            rows.append(payload)

        if rows:
            _insert_sheet_rows(conn, APP_KIND, sheet_name, rows)
            conn.commit()


def load_app_master_df() -> pd.DataFrame:
    if not postgres_configured() or psycopg2 is None:
        df = _load_workbook_sheet_df(_APP_FALLBACK_PATH, MASTER_SHEET, header=1)
        if df.empty:
            return df
        if LANG_COL in df.columns:
            df[LANG_COL] = df[LANG_COL].ffill()
        return df
    df = load_app_sheet(MASTER_SHEET)
    if df.empty:
        return df
    if LANG_COL in df.columns:
        df[LANG_COL] = df[LANG_COL].ffill()
    return df


def load_db_records() -> list[dict[str, str]]:
    df = load_app_master_df()
    if df.empty or LANG_COL not in df.columns or URL_COL not in df.columns:
        return []

    out: list[dict[str, str]] = []
    for _, row in df.iterrows():
        lang = str(row.get(LANG_COL, "") or "").strip()
        url = str(row.get(URL_COL, "") or "").strip()
        if not lang or not url:
            continue
        if _clean_url(url) in JUNK_VALUES:
            continue
        out.append({"language": lang, "url": url})
    return out


def list_db_languages() -> list[str]:
    df = load_app_master_df()
    if df.empty or LANG_COL not in df.columns:
        return []
    seen: set[str] = set()
    ordered: list[str] = []
    for value in df[LANG_COL].dropna().astype(str).tolist():
        item = value.strip()
        if not item:
            continue
        key = _normalize_text(item)
        if key in seen:
            continue
        seen.add(key)
        ordered.append(item)
    return ordered


def add_language_to_db(language: str) -> bool:
    language = str(language or "").strip()
    if not language:
        return False

    if not postgres_configured() or psycopg2 is None:
        return False

    with _connect() as conn:
        df = _load_sheet_df(conn, APP_KIND, MASTER_SHEET)
        if not df.empty and LANG_COL in df.columns:
            existing = {
                _normalize_text(v)
                for v in df[LANG_COL].dropna().astype(str).tolist()
                if str(v).strip()
            }
            if _normalize_text(language) in existing:
                return False

        columns = list(df.columns) if not df.empty else ["ID", URL_COL, LANG_COL]
        payload = {col: None for col in columns}
        if "ID" in payload:
            id_values = pd.to_numeric(df["ID"], errors="coerce").dropna() if "ID" in df.columns else pd.Series(dtype=float)
            payload["ID"] = int(id_values.max()) + 1 if not id_values.empty else 1
        if URL_COL in payload:
            payload[URL_COL] = ""
        if LANG_COL in payload:
            payload[LANG_COL] = language
        _insert_sheet_rows(conn, APP_KIND, MASTER_SHEET, [payload])
        conn.commit()
    return True


def remove_language_from_db(language: str) -> bool:
    language = str(language or "").strip()
    if not language:
        return False

    if not postgres_configured() or psycopg2 is None:
        return False

    with _connect() as conn:
        df = load_app_master_df()
        if df.empty or LANG_COL not in df.columns:
            return False

        target = _normalize_text(language)
        existing = {
            _normalize_text(v)
            for v in df[LANG_COL].dropna().astype(str).tolist()
            if str(v).strip()
        }
        if target not in existing:
            return False

        with conn.cursor() as cur:
            cur.execute(
                """
                DELETE FROM reference_rows
                WHERE kind = %s
                  AND sheet_name = %s
                  AND lower(trim(coalesce(payload ->> %s, ''))) = %s
                """,
                (APP_KIND, MASTER_SHEET, LANG_COL, target),
            )
        conn.commit()
    return True


def get_city_sheet_names() -> list[str]:
    if not postgres_configured() or psycopg2 is None:
        return _workbook_sheet_names(_CITY_FALLBACK_PATH, exclude_summary=True)
    names = list_sheet_names(CITY_KIND)
    return [name for name in names if name.strip().lower() != "summary by state"]


def fallback_city_sheet_names_from_workbook(excel_path: str | Path) -> list[str]:
    return _workbook_sheet_names(excel_path, exclude_summary=True)


def load_city_db_sheet(sheet_names_str: str) -> list[dict]:
    if not sheet_names_str:
        return []

    sheet_names = [sn.strip() for sn in sheet_names_str.split(",") if sn.strip()]
    seen_cities: dict[str, dict[str, Any]] = {}

    for sn in sheet_names:
        df = load_sheet_df(CITY_KIND, sn)
        if df.empty or "City" not in df.columns:
            continue

        weight_col = next(
            (c for c in ["Potential Impressions", "Unique Cookies w/ Impressions"] if c in df.columns),
            None,
        )
        creative_col = next((c for c in df.columns if "creative" in c.lower()), None)

        df = df[df["City"].notna()].copy()
        for _, row in df.iterrows():
            city_name = str(row["City"]).strip()
            if not city_name or city_name.lower() in {"nan", "none"}:
                continue
            weight = _safe_float(row[weight_col], 1.0) if weight_col else 1.0
            creatives: list[str] = []
            if creative_col and pd.notna(row.get(creative_col)):
                raw_c = str(row[creative_col])
                creatives = [c.strip() for c in raw_c.replace("|", ",").split(",") if c.strip()]

            key = city_name.lower()
            if key in seen_cities:
                seen_cities[key]["weight"] += weight
                seen_cities[key]["creatives"].update(creatives)
            else:
                seen_cities[key] = {
                    "name": city_name,
                    "weight": weight,
                    "creatives": set(creatives),
                }

    merged = [
        {
            "name": v["name"],
            "weight": float(v["weight"]),
            "creatives": sorted(list(v["creatives"])),
        }
        for v in seen_cities.values()
    ]
    return sorted(merged, key=lambda x: x["weight"], reverse=True)


def load_city_reference(n_cities: int = 50) -> list[tuple[str, float]]:
    df = load_sheet_df(CITY_KIND, MASTER_SHEET)
    if df.empty or "City" not in df.columns:
        return []
    weight_col = next(
        (c for c in ["Potential Impressions", "Unique Cookies w/ Impressions"] if c in df.columns),
        None,
    )
    df = df[df["City"].notna()].copy()
    df["_city"] = df["City"].astype(str).str.strip()
    df = df[df["_city"].str.lower().notnull() & (df["_city"] != "")]
    if weight_col:
        df["_weight"] = df[weight_col].apply(lambda v: _safe_float(v, 1.0))
    else:
        df["_weight"] = 1.0
    df = df[df["_weight"] > 0].sort_values("_weight", ascending=False)
    return [(row["_city"], float(row["_weight"])) for _, row in df.head(n_cities).iterrows()]
