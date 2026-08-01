from __future__ import annotations

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from catalog import database_options, filter_by_option, latest_record, numeric_options
from database import (
    add_item,
    connect,
    fetch_inventory_dataframe,
    fetch_movements_dataframe,
    get_item,
    initialize_database,
    insert_items,
    make_inventory_code,
    recode_all_inventory,
    record_stock_movement,
    update_item,
)
from importer import parse_source_workbook
from reports import build_report_frame, create_excel_report, inventory_summary


def main() -> None:
    app_dir = Path(__file__).resolve().parent
    seed = app_dir / "data" / "Inventory_list_source.xlsx"
    assert seed.exists(), "Seed workbook is missing"

    records, diagnostics = parse_source_workbook(seed)
    assert diagnostics["records"] == 407, diagnostics
    assert len(records) == 407

    with TemporaryDirectory() as temporary_directory:
        database_path = Path(temporary_directory) / "inventory_test.db"
        initialize_database(database_path)
        result = insert_items(records, replace=True, db_path=database_path)
        assert result == {"inserted": 407, "skipped": 0}, result

        inventory = fetch_inventory_dataframe(db_path=database_path, warning_days=30)
        assert len(inventory) == 407
        assert inventory["inventory_code"].is_unique
        assert sorted(inventory["id"].tolist()) == list(range(1, 408))
        assert all(
            row.inventory_code == make_inventory_code(row.item_name, row.description, row.id)
            for row in inventory.itertuples()
        )
        assert inventory["source_row"].min() == 4
        assert inventory["source_row"].max() == 410

        item_options = database_options(inventory, "item_name", include_missing=False)
        assert len(item_options) == 29
        selected_item = "SGOT"
        item_rows = filter_by_option(inventory, "item_name", selected_item)
        assert not item_rows.empty
        brand_options = database_options(item_rows, "brand")
        brand_rows = filter_by_option(item_rows, "brand", brand_options[0])
        assert not brand_rows.empty
        assert latest_record(brand_rows) is not None
        assert numeric_options(brand_rows, "unit_price")

        with connect(database_path) as connection:
            connection.execute(
                "UPDATE inventory SET inventory_code = printf('INV-%05d', id)"
            )
            connection.commit()
        assert recode_all_inventory(database_path) == 407
        migrated = get_item(1, database_path)
        assert migrated["inventory_code"] == make_inventory_code(
            migrated["item_name"], migrated["description"], 1
        )

        test_item_id = add_item(
            {
                "item_name": "SMOKE TEST ITEM",
                "description": "2 x 30 ML",
                "brand": "TEST",
                "quantity_in_stock": None,
                "reorder_level": 2,
                "unit_price": 10,
                "expiry_date": "2030-12-31",
            },
            db_path=database_path,
        )
        assert test_item_id == 408
        created = get_item(test_item_id, database_path)
        assert created["inventory_code"] == "STI-2X30ML-00408"
        update_item(test_item_id, {"description": "5X100"}, database_path)
        created = get_item(test_item_id, database_path)
        assert created["inventory_code"] == "STI-5X100-00408"
        balance = record_stock_movement(
            test_item_id,
            "Receive",
            5,
            reference="TEST-RECEIPT",
            db_path=database_path,
        )
        assert balance == 5
        balance = record_stock_movement(
            test_item_id,
            "Issue",
            2,
            reference="TEST-ISSUE",
            db_path=database_path,
        )
        assert balance == 3
        assert get_item(test_item_id, database_path)["quantity_in_stock"] == 3
        assert len(fetch_movements_dataframe(database_path)) == 2

        inventory = fetch_inventory_dataframe(db_path=database_path, warning_days=30)
        report = build_report_frame("Full Inventory Report", inventory)
        report_bytes = create_excel_report(
            "Full Inventory Report",
            report,
            inventory_summary(inventory),
            "₹",
        )
        workbook = load_workbook(BytesIO(report_bytes), read_only=True)
        assert workbook.sheetnames == ["Summary", "Report Data"]
        assert workbook["Report Data"].max_row == len(report) + 1
        assert workbook["Report Data"]["A1"].value == "Inventory ID"

    print(
        "Smoke test passed: 407 source records imported; CRUD, cascading database options, "
        "incremental IDs, code migration, stock movements, "
        "calculated statuses and Excel reporting verified."
    )


if __name__ == "__main__":
    main()
