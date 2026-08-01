from __future__ import annotations

import math
import os
import re
import sqlite3
import tempfile
import uuid
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Iterator, Mapping

import pandas as pd


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DEFAULT_DB_PATH = DATA_DIR / "inventory.db"


INVENTORY_FIELDS = (
    "inventory_code",
    "item_name",
    "brand",
    "category",
    "test_volume",
    "description",
    "unit_price",
    "total_test_effective",
    "total_test_available",
    "total_test_done",
    "selling_price_per_test",
    "quantity_in_stock",
    "reorder_level",
    "reorder_quantity",
    "create_date",
    "opened_date",
    "expiry_date",
    "finish_date",
    "reminder_days",
    "reorder_due_date",
    "batch_no",
    "supplier",
    "storage_location",
    "notes",
    "source_inventory_id",
    "source_row",
    "source_notes",
    "is_active",
)


def _code_tokens(value: Any) -> list[str]:
    return re.findall(r"[A-Z0-9]+", str(value or "").upper())


def item_short_form(item_name: Any) -> str:
    """Return a stable, readable short form for an inventory item name."""
    tokens = _code_tokens(item_name)
    if not tokens:
        return "ITEM"
    if len(tokens) == 1:
        token = tokens[0]
        return token if len(token) <= 6 else token[:4]
    if any(character.isdigit() for character in tokens[0]):
        short = tokens[0][:6] + "".join(token[0] for token in tokens[1:] if token)
    else:
        short = "".join(token[0] for token in tokens if token)
    return short[:8] or "ITEM"


def description_short_form(description: Any) -> str:
    """Return a compact pack/description component for an inventory code."""
    tokens = _code_tokens(description)
    if not tokens:
        return "NA"
    compact = "".join(tokens)
    return compact[:12]


def inventory_code_base(item_name: Any, description: Any) -> str:
    return f"{item_short_form(item_name)}-{description_short_form(description)}"


def make_inventory_code(item_name: Any, description: Any, item_id: int) -> str:
    """Build ITEMSHORT-DESCRIPTION-ID, with ID keeping historical lots unique."""
    return f"{inventory_code_base(item_name, description)}-{int(item_id):05d}"


def _temporary_inventory_code() -> str:
    return f"TMP-{uuid.uuid4().hex.upper()}"


def resolve_db_path(db_path: str | Path | None = None) -> Path:
    if db_path is not None:
        return Path(db_path)
    configured = os.getenv("INVENTORY_DB_PATH")
    return Path(configured) if configured else DEFAULT_DB_PATH


def connect(db_path: str | Path | None = None) -> sqlite3.Connection:
    path = resolve_db_path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA journal_mode = WAL")
    connection.execute("PRAGMA busy_timeout = 30000")
    return connection


@contextmanager
def transaction(db_path: str | Path | None = None) -> Iterator[sqlite3.Connection]:
    connection = connect(db_path)
    try:
        connection.execute("BEGIN IMMEDIATE")
        yield connection
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _migrate_inventory_schema(connection: sqlite3.Connection) -> None:
    """Add fields introduced after the original database without rewriting old rows."""
    columns = {
        row["name"] for row in connection.execute("PRAGMA table_info(inventory)").fetchall()
    }
    additions = {
        "create_date": "TEXT",
        "total_test_available": (
            "REAL CHECK (total_test_available IS NULL OR total_test_available >= 0)"
        ),
        "total_test_done": "REAL CHECK (total_test_done IS NULL OR total_test_done >= 0)",
        "finish_date": "TEXT",
        "deleted_at": "TEXT",
    }
    for column, definition in additions.items():
        if column not in columns:
            connection.execute(f"ALTER TABLE inventory ADD COLUMN {column} {definition}")

    # Preserve records archived by earlier releases as soft-deleted records.
    if "archived_at" in columns:
        connection.execute(
            """
            UPDATE inventory
            SET deleted_at = COALESCE(deleted_at, archived_at)
            WHERE is_active = 0
            """
        )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_inventory_deleted ON inventory(deleted_at)"
    )


def initialize_database(db_path: str | Path | None = None) -> None:
    schema = """
    CREATE TABLE IF NOT EXISTS inventory (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        inventory_code TEXT NOT NULL UNIQUE,
        item_name TEXT NOT NULL,
        brand TEXT,
        category TEXT NOT NULL DEFAULT 'Laboratory Inventory',
        test_volume TEXT,
        description TEXT,
        unit_price REAL CHECK (unit_price IS NULL OR unit_price >= 0),
        total_test_effective REAL CHECK (total_test_effective IS NULL OR total_test_effective >= 0),
        total_test_available REAL CHECK (total_test_available IS NULL OR total_test_available >= 0),
        total_test_done REAL CHECK (total_test_done IS NULL OR total_test_done >= 0),
        selling_price_per_test REAL CHECK (selling_price_per_test IS NULL OR selling_price_per_test >= 0),
        quantity_in_stock REAL CHECK (quantity_in_stock IS NULL OR quantity_in_stock >= 0),
        reorder_level REAL CHECK (reorder_level IS NULL OR reorder_level >= 0),
        reorder_quantity REAL CHECK (reorder_quantity IS NULL OR reorder_quantity >= 0),
        create_date TEXT,
        opened_date TEXT,
        expiry_date TEXT,
        finish_date TEXT,
        reminder_days INTEGER CHECK (reminder_days IS NULL OR reminder_days >= 0),
        reorder_due_date TEXT,
        batch_no TEXT,
        supplier TEXT,
        storage_location TEXT,
        notes TEXT,
        source_inventory_id TEXT,
        source_row INTEGER,
        source_notes TEXT,
        is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        deleted_at TEXT
    );

    CREATE TABLE IF NOT EXISTS stock_movements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        inventory_id INTEGER NOT NULL,
        movement_type TEXT NOT NULL CHECK (movement_type IN ('Receive', 'Issue', 'Adjustment')),
        quantity_change REAL NOT NULL,
        quantity_before REAL NOT NULL,
        quantity_after REAL NOT NULL CHECK (quantity_after >= 0),
        reference TEXT,
        notes TEXT,
        moved_at TEXT NOT NULL,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (inventory_id) REFERENCES inventory(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS settings (
        setting_key TEXT PRIMARY KEY,
        setting_value TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE INDEX IF NOT EXISTS idx_inventory_item ON inventory(item_name);
    CREATE INDEX IF NOT EXISTS idx_inventory_brand ON inventory(brand);
    CREATE INDEX IF NOT EXISTS idx_inventory_expiry ON inventory(expiry_date);
    CREATE INDEX IF NOT EXISTS idx_inventory_active ON inventory(is_active);
    CREATE INDEX IF NOT EXISTS idx_movements_inventory ON stock_movements(inventory_id);
    CREATE INDEX IF NOT EXISTS idx_movements_date ON stock_movements(moved_at);
    """
    defaults = {
        "business_name": "Inventory Control Centre",
        "currency_symbol": "₹",
        "expiry_warning_days": "30",
    }
    with connect(db_path) as connection:
        connection.executescript(schema)
        _migrate_inventory_schema(connection)
        connection.executemany(
            "INSERT OR IGNORE INTO settings(setting_key, setting_value) VALUES (?, ?)",
            defaults.items(),
        )
        connection.commit()
    recode_all_inventory(db_path)


def get_settings(db_path: str | Path | None = None) -> dict[str, str]:
    with connect(db_path) as connection:
        rows = connection.execute(
            "SELECT setting_key, setting_value FROM settings ORDER BY setting_key"
        ).fetchall()
    return {row["setting_key"]: row["setting_value"] for row in rows}


def save_settings(values: Mapping[str, Any], db_path: str | Path | None = None) -> None:
    with transaction(db_path) as connection:
        connection.executemany(
            """
            INSERT INTO settings(setting_key, setting_value, updated_at)
            VALUES (?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(setting_key) DO UPDATE SET
                setting_value = excluded.setting_value,
                updated_at = CURRENT_TIMESTAMP
            """,
            [(str(key), str(value)) for key, value in values.items()],
        )


def inventory_count(db_path: str | Path | None = None) -> int:
    with connect(db_path) as connection:
        return int(connection.execute("SELECT COUNT(*) FROM inventory").fetchone()[0])


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def clean_value(value: Any) -> Any:
    if _is_missing(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date().isoformat()
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if hasattr(value, "item"):
        try:
            return value.item()
        except (ValueError, AttributeError):
            pass
    return value


def _normalise_item(item: Mapping[str, Any]) -> dict[str, Any]:
    normalised = {field: clean_value(item.get(field)) for field in INVENTORY_FIELDS}
    item_name = str(normalised.get("item_name") or "").strip()
    if not item_name:
        raise ValueError("Item name is required.")
    normalised["item_name"] = item_name
    normalised["inventory_code"] = str(normalised.get("inventory_code") or "").strip().upper() or None
    normalised["brand"] = _clean_text(normalised.get("brand"))
    normalised["category"] = _clean_text(normalised.get("category")) or "Laboratory Inventory"
    normalised["is_active"] = 1 if normalised.get("is_active") in (None, True, 1, "1", "Yes") else 0
    for field in (
        "unit_price",
        "total_test_effective",
        "total_test_available",
        "total_test_done",
        "selling_price_per_test",
        "quantity_in_stock",
        "reorder_level",
        "reorder_quantity",
        "reminder_days",
    ):
        value = normalised.get(field)
        if value is not None and float(value) < 0:
            raise ValueError(f"{field.replace('_', ' ').title()} cannot be negative.")
    return normalised


def _clean_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def recode_all_inventory(db_path: str | Path | None = None) -> int:
    """Migrate every record to the current ITEMSHORT-DESCRIPTION-ID convention."""
    with transaction(db_path) as connection:
        rows = connection.execute(
            "SELECT id, item_name, description, inventory_code FROM inventory ORDER BY id"
        ).fetchall()
        changes = [
            (
                int(row["id"]),
                make_inventory_code(row["item_name"], row["description"], row["id"]),
            )
            for row in rows
            if row["inventory_code"]
            != make_inventory_code(row["item_name"], row["description"], row["id"])
        ]
        # A temporary first pass prevents unique-key collisions with any legacy
        # code that happens to equal another record's new code.
        for item_id, _ in changes:
            connection.execute(
                "UPDATE inventory SET inventory_code = ? WHERE id = ?",
                (_temporary_inventory_code(), item_id),
            )
        for item_id, new_code in changes:
            connection.execute(
                "UPDATE inventory SET inventory_code = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                (new_code, item_id),
            )
    return len(changes)


def insert_items(
    items: Iterable[Mapping[str, Any]],
    *,
    replace: bool = False,
    db_path: str | Path | None = None,
) -> dict[str, int]:
    inserted = 0
    skipped = 0
    with transaction(db_path) as connection:
        if replace:
            connection.execute("DELETE FROM stock_movements")
            connection.execute("DELETE FROM inventory")
        for raw_item in items:
            try:
                item = _normalise_item(raw_item)
            except (TypeError, ValueError):
                skipped += 1
                continue
            item["inventory_code"] = _temporary_inventory_code()
            columns = [field for field in INVENTORY_FIELDS]
            placeholders = ", ".join("?" for _ in columns)
            cursor = connection.execute(
                f"INSERT INTO inventory ({', '.join(columns)}) VALUES ({placeholders})",
                [item.get(column) for column in columns],
            )
            item_id = int(cursor.lastrowid)
            connection.execute(
                "UPDATE inventory SET inventory_code = ? WHERE id = ?",
                (make_inventory_code(item["item_name"], item.get("description"), item_id), item_id),
            )
            inserted += 1
    return {"inserted": inserted, "skipped": skipped}


def add_item(item: Mapping[str, Any], db_path: str | Path | None = None) -> int:
    normalised = _normalise_item(item)
    normalised["create_date"] = date.today().isoformat()
    with transaction(db_path) as connection:
        normalised["inventory_code"] = _temporary_inventory_code()
        columns = [field for field in INVENTORY_FIELDS]
        placeholders = ", ".join("?" for _ in columns)
        cursor = connection.execute(
            f"INSERT INTO inventory ({', '.join(columns)}) VALUES ({placeholders})",
            [normalised.get(column) for column in columns],
        )
        item_id = int(cursor.lastrowid)
        inventory_code = make_inventory_code(
            normalised["item_name"], normalised.get("description"), item_id
        )
        connection.execute(
            "UPDATE inventory SET inventory_code = ? WHERE id = ?",
            (inventory_code, item_id),
        )
        return item_id


def update_item(
    item_id: int,
    updates: Mapping[str, Any],
    db_path: str | Path | None = None,
) -> None:
    allowed = set(INVENTORY_FIELDS) - {"source_row", "inventory_code", "create_date"}
    values: dict[str, Any] = {}
    for key, value in updates.items():
        if key not in allowed:
            continue
        values[key] = clean_value(value)
    if "item_name" in values:
        values["item_name"] = str(values["item_name"] or "").strip()
        if not values["item_name"]:
            raise ValueError("Item name is required.")
    for field in (
        "unit_price",
        "total_test_effective",
        "total_test_available",
        "total_test_done",
        "selling_price_per_test",
        "quantity_in_stock",
        "reorder_level",
        "reorder_quantity",
        "reminder_days",
    ):
        if field in values and values[field] is not None and float(values[field]) < 0:
            raise ValueError(f"{field.replace('_', ' ').title()} cannot be negative.")
    if not values:
        return
    assignments = ", ".join(f"{column} = ?" for column in values)
    with transaction(db_path) as connection:
        cursor = connection.execute(
            f"UPDATE inventory SET {assignments}, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            [*values.values(), int(item_id)],
        )
        if cursor.rowcount != 1:
            raise ValueError("Inventory record was not found.")
        row = connection.execute(
            "SELECT item_name, description FROM inventory WHERE id = ?", (int(item_id),)
        ).fetchone()
        connection.execute(
            "UPDATE inventory SET inventory_code = ? WHERE id = ?",
            (make_inventory_code(row["item_name"], row["description"], item_id), int(item_id)),
        )


def bulk_update_items(
    rows: Iterable[Mapping[str, Any]],
    db_path: str | Path | None = None,
) -> int:
    allowed = {
        "unit_price",
        "selling_price_per_test",
        "quantity_in_stock",
        "reorder_level",
        "reorder_quantity",
        "total_test_available",
        "total_test_done",
        "expiry_date",
        "finish_date",
    }
    updated = 0
    with transaction(db_path) as connection:
        for row in rows:
            if "id" not in row:
                continue
            values = {key: clean_value(value) for key, value in row.items() if key in allowed}
            for field in allowed - {"expiry_date", "finish_date"}:
                if field in values and values[field] is not None and float(values[field]) < 0:
                    raise ValueError(f"{field.replace('_', ' ').title()} cannot be negative.")
            if not values:
                continue
            assignments = ", ".join(f"{column} = ?" for column in values)
            cursor = connection.execute(
                f"UPDATE inventory SET {assignments}, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                [*values.values(), int(row["id"])],
            )
            updated += cursor.rowcount
    return updated


def delete_item(item_id: int, db_path: str | Path | None = None) -> None:
    with transaction(db_path) as connection:
        cursor = connection.execute(
            """
            UPDATE inventory
            SET is_active = 0, deleted_at = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (
                datetime.now().isoformat(timespec="seconds"),
                int(item_id),
            ),
        )
        if cursor.rowcount != 1:
            raise ValueError("Inventory record was not found.")


def restore_item(item_id: int, db_path: str | Path | None = None) -> None:
    with transaction(db_path) as connection:
        cursor = connection.execute(
            """
            UPDATE inventory
            SET is_active = 1, deleted_at = NULL, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (int(item_id),),
        )
        if cursor.rowcount != 1:
            raise ValueError("Inventory record was not found.")


def get_item(item_id: int, db_path: str | Path | None = None) -> dict[str, Any] | None:
    with connect(db_path) as connection:
        row = connection.execute("SELECT * FROM inventory WHERE id = ?", (int(item_id),)).fetchone()
    return dict(row) if row else None


def fetch_inventory_dataframe(
    *,
    include_deleted: bool = False,
    deleted_only: bool = False,
    warning_days: int = 30,
    db_path: str | Path | None = None,
) -> pd.DataFrame:
    if deleted_only:
        where = "WHERE is_active = 0"
    elif include_deleted:
        where = ""
    else:
        where = "WHERE is_active = 1"
    with connect(db_path) as connection:
        frame = pd.read_sql_query(
            f"SELECT * FROM inventory {where} ORDER BY item_name, expiry_date, inventory_code",
            connection,
        )
    if frame.empty:
        return _add_calculated_columns(frame, warning_days)
    for column in (
        "create_date",
        "opened_date",
        "expiry_date",
        "finish_date",
        "reorder_due_date",
        "deleted_at",
    ):
        frame[column] = pd.to_datetime(frame[column], errors="coerce")
    return _add_calculated_columns(frame, warning_days)


def _add_calculated_columns(frame: pd.DataFrame, warning_days: int) -> pd.DataFrame:
    frame = frame.copy()
    if frame.empty:
        for column in (
            "cost_per_test",
            "profit_per_test",
            "inventory_value",
            "days_to_expiry",
            "stock_status",
            "expiry_status",
            "overall_status",
            "data_quality_issues",
        ):
            frame[column] = pd.Series(dtype="object")
        return frame

    frame["cost_per_test"] = frame["unit_price"].div(
        frame["total_test_effective"].where(frame["total_test_effective"].gt(0))
    )
    frame["profit_per_test"] = frame["selling_price_per_test"] - frame["cost_per_test"]
    frame["inventory_value"] = frame["unit_price"] * frame["quantity_in_stock"]
    today = pd.Timestamp(date.today())
    frame["days_to_expiry"] = (frame["expiry_date"] - today).dt.days

    def stock_status(row: pd.Series) -> str:
        quantity = row.get("quantity_in_stock")
        reorder_level = row.get("reorder_level")
        if pd.isna(quantity):
            return "Stock Not Set"
        if float(quantity) <= 0:
            return "Out of Stock"
        if not pd.isna(reorder_level) and float(quantity) <= float(reorder_level):
            return "Low Stock"
        return "In Stock"

    def expiry_status(row: pd.Series) -> str:
        days = row.get("days_to_expiry")
        if pd.isna(days):
            return "Expiry Not Set"
        if int(days) < 0:
            return "Expired"
        if int(days) <= int(warning_days):
            return "Expiring Soon"
        return "Valid"

    def overall_status(row: pd.Series) -> str:
        if int(row.get("is_active", 1)) == 0:
            return "Deleted"
        if row["expiry_status"] in ("Expired", "Expiring Soon"):
            return row["expiry_status"]
        return row["stock_status"]

    def quality_issues(row: pd.Series) -> str:
        issues: list[str] = []
        if pd.isna(row.get("quantity_in_stock")):
            issues.append("Stock quantity missing")
        if pd.isna(row.get("expiry_date")):
            issues.append("Expiry date missing")
        if pd.isna(row.get("unit_price")):
            issues.append("Unit price missing")
        if pd.isna(row.get("reorder_level")):
            issues.append("Reorder level missing")
        if str(row.get("source_notes") or "").strip():
            issues.append("Source value needs review")
        return "; ".join(issues)

    frame["stock_status"] = frame.apply(stock_status, axis=1)
    frame["expiry_status"] = frame.apply(expiry_status, axis=1)
    frame["overall_status"] = frame.apply(overall_status, axis=1)
    frame["data_quality_issues"] = frame.apply(quality_issues, axis=1)
    return frame


def record_stock_movement(
    item_id: int,
    movement_type: str,
    quantity: float,
    *,
    reference: str | None = None,
    notes: str | None = None,
    moved_at: date | datetime | str | None = None,
    db_path: str | Path | None = None,
) -> float:
    movement_type = movement_type.title()
    if movement_type not in {"Receive", "Issue", "Adjustment"}:
        raise ValueError("Movement type must be Receive, Issue, or Adjustment.")
    quantity = float(quantity)
    if quantity < 0:
        raise ValueError("Quantity cannot be negative.")
    movement_date = clean_value(moved_at) or date.today().isoformat()
    with transaction(db_path) as connection:
        row = connection.execute(
            "SELECT quantity_in_stock FROM inventory WHERE id = ? AND is_active = 1",
            (int(item_id),),
        ).fetchone()
        if row is None:
            raise ValueError("Active inventory record was not found.")
        current = 0.0 if row[0] is None else float(row[0])
        if movement_type == "Receive":
            after = current + quantity
            change = quantity
        elif movement_type == "Issue":
            if quantity > current:
                raise ValueError(f"Issue quantity exceeds available stock ({current:g}).")
            after = current - quantity
            change = -quantity
        else:
            after = quantity
            change = after - current
        connection.execute(
            "UPDATE inventory SET quantity_in_stock = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            (after, int(item_id)),
        )
        connection.execute(
            """
            INSERT INTO stock_movements(
                inventory_id, movement_type, quantity_change, quantity_before,
                quantity_after, reference, notes, moved_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(item_id),
                movement_type,
                change,
                current,
                after,
                _clean_text(reference),
                _clean_text(notes),
                movement_date,
            ),
        )
    return after


def fetch_movements_dataframe(db_path: str | Path | None = None) -> pd.DataFrame:
    query = """
        SELECT
            m.id,
            m.moved_at,
            i.id AS inventory_id,
            i.inventory_code,
            i.item_name,
            i.brand,
            m.movement_type,
            m.quantity_change,
            m.quantity_before,
            m.quantity_after,
            m.reference,
            m.notes,
            m.created_at
        FROM stock_movements m
        JOIN inventory i ON i.id = m.inventory_id
        ORDER BY m.moved_at DESC, m.id DESC
    """
    with connect(db_path) as connection:
        frame = pd.read_sql_query(query, connection)
    if not frame.empty:
        frame["moved_at"] = pd.to_datetime(frame["moved_at"], errors="coerce")
        frame["created_at"] = pd.to_datetime(frame["created_at"], errors="coerce")
    return frame


def create_database_backup(db_path: str | Path | None = None) -> bytes:
    source = connect(db_path)
    try:
        with tempfile.NamedTemporaryFile(suffix=".db") as temp_file:
            destination = sqlite3.connect(temp_file.name)
            try:
                source.backup(destination)
            finally:
                destination.close()
            return Path(temp_file.name).read_bytes()
    finally:
        source.close()
