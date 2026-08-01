from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from database import (
    add_item,
    archive_item,
    bulk_update_items,
    create_database_backup,
    fetch_inventory_dataframe,
    fetch_movements_dataframe,
    get_settings,
    initialize_database,
    insert_items,
    inventory_count,
    record_stock_movement,
    save_settings,
    update_item,
)
from importer import parse_source_workbook, parse_tabular_upload, records_preview
from reports import (
    build_report_frame,
    create_excel_report,
    create_import_template,
    csv_bytes,
    display_frame,
    inventory_summary,
)


APP_DIR = Path(__file__).resolve().parent
SEED_WORKBOOK = APP_DIR / "data" / "Inventory_list_source.xlsx"


st.set_page_config(
    page_title="Inventory Control Centre",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)


CUSTOM_CSS = """
<style>
    :root {
        --navy: #12304a;
        --teal: #00a6a6;
        --soft: #f4f8fb;
        --line: #d9e4ec;
    }
    .stApp { background: #f7fafc; }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #102b42 0%, #173f5f 100%);
    }
    [data-testid="stSidebar"] * { color: #f7fbff; }
    [data-testid="stSidebar"] div[role="radiogroup"] label {
        padding: .55rem .7rem;
        border-radius: .55rem;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background: rgba(255,255,255,.09);
    }
    h1, h2, h3 { color: var(--navy); letter-spacing: -0.02em; }
    .page-kicker {
        color: var(--teal);
        font-size: .78rem;
        font-weight: 800;
        letter-spacing: .12em;
        text-transform: uppercase;
        margin-bottom: .2rem;
    }
    .page-subtitle { color: #637381; margin-top: -.55rem; margin-bottom: 1.2rem; }
    [data-testid="stMetric"] {
        background: white;
        border: 1px solid var(--line);
        border-radius: .8rem;
        padding: .85rem 1rem;
        box-shadow: 0 4px 14px rgba(20, 48, 74, .05);
    }
    [data-testid="stMetricLabel"] { color: #637381; }
    [data-testid="stMetricValue"] { color: var(--navy); }
    .status-note {
        background: #e8f7f7;
        border-left: 4px solid var(--teal);
        color: var(--navy);
        padding: .75rem 1rem;
        border-radius: .4rem;
        margin-bottom: 1rem;
    }
    div[data-testid="stDataFrame"] {
        border: 1px solid var(--line);
        border-radius: .7rem;
        overflow: hidden;
    }
    .block-container { padding-top: 1.8rem; padding-bottom: 3rem; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def page_header(title: str, subtitle: str) -> None:
    st.markdown('<div class="page-kicker">Inventory Management</div>', unsafe_allow_html=True)
    st.title(title)
    st.markdown(f'<div class="page-subtitle">{subtitle}</div>', unsafe_allow_html=True)


def money(value: Any, symbol: str) -> str:
    try:
        if value is None or pd.isna(value):
            return "—"
        return f"{symbol}{float(value):,.2f}"
    except (TypeError, ValueError):
        return "—"


def optional_number(text: str, label: str) -> float | None:
    clean = str(text or "").strip().replace(",", "")
    if not clean:
        return None
    try:
        value = float(clean)
    except ValueError as exc:
        raise ValueError(f"{label} must be a number or left blank.") from exc
    if value < 0:
        raise ValueError(f"{label} cannot be negative.")
    return value


def numeric_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
        number = float(value)
        return f"{number:g}"
    except (TypeError, ValueError):
        return str(value)


def text_value(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value)


def date_value(value: Any) -> date:
    try:
        if value is not None and not pd.isna(value):
            return pd.Timestamp(value).date()
    except (TypeError, ValueError):
        pass
    return date.today()


def apply_inventory_filters(frame: pd.DataFrame, key_prefix: str) -> pd.DataFrame:
    if frame.empty:
        return frame
    row1 = st.columns([2.1, 1.2, 1.2, 1.2])
    search = row1[0].text_input(
        "Search",
        placeholder="Item, code, brand, batch or description",
        key=f"{key_prefix}_search",
    )
    brands = sorted(value for value in frame["brand"].dropna().astype(str).unique() if value)
    selected_brands = row1[1].multiselect("Brand", brands, key=f"{key_prefix}_brand")
    stock_values = sorted(frame["stock_status"].dropna().astype(str).unique())
    selected_stock = row1[2].multiselect(
        "Stock status", stock_values, key=f"{key_prefix}_stock"
    )
    expiry_values = sorted(frame["expiry_status"].dropna().astype(str).unique())
    selected_expiry = row1[3].multiselect(
        "Expiry status", expiry_values, key=f"{key_prefix}_expiry"
    )
    filtered = frame.copy()
    if search.strip():
        pattern = re.escape(search.strip())
        searchable_columns = [
            column
            for column in (
                "inventory_code",
                "item_name",
                "brand",
                "description",
                "batch_no",
                "supplier",
            )
            if column in filtered.columns
        ]
        combined = filtered[searchable_columns].fillna("").astype(str).agg(" ".join, axis=1)
        filtered = filtered.loc[combined.str.contains(pattern, case=False, regex=True)]
    if selected_brands:
        filtered = filtered.loc[filtered["brand"].isin(selected_brands)]
    if selected_stock:
        filtered = filtered.loc[filtered["stock_status"].isin(selected_stock)]
    if selected_expiry:
        filtered = filtered.loc[filtered["expiry_status"].isin(selected_expiry)]
    return filtered


def initialize_app_data() -> None:
    initialize_database()
    if inventory_count() == 0 and SEED_WORKBOOK.exists():
        records, diagnostics = parse_source_workbook(SEED_WORKBOOK)
        result = insert_items(records)
        st.session_state["seed_result"] = {
            **result,
            **diagnostics,
        }


def render_dashboard(inventory: pd.DataFrame, currency: str) -> None:
    page_header(
        "Dashboard",
        "Live stock, expiry, reorder and data-quality overview from the SQLite database.",
    )
    summary = inventory_summary(inventory)
    columns = st.columns(6)
    columns[0].metric("Inventory Records", f"{summary['Total Records']:,}")
    columns[1].metric("Unique Items", f"{summary['Unique Items']:,}")
    columns[2].metric("Known Stock Value", money(summary["Known Inventory Value"], currency))
    columns[3].metric("Reorder Alerts", f"{summary['Reorder Alerts']:,}")
    columns[4].metric("Expired", f"{summary['Expired Records']:,}")
    columns[5].metric("Expiring Soon", f"{summary['Expiring Soon']:,}")

    if inventory.empty:
        st.info("No active inventory records are available yet.")
        return

    stock_known = int(inventory["quantity_in_stock"].notna().sum())
    expiry_known = int(inventory["expiry_date"].notna().sum())
    price_known = int(inventory["unit_price"].notna().sum())
    reorder_known = int(inventory["reorder_level"].notna().sum())
    completeness = (stock_known + expiry_known + price_known + reorder_known) / (len(inventory) * 4)
    st.markdown(
        f'<div class="status-note"><b>Data readiness:</b> {completeness:.0%} complete across stock, expiry, price and reorder fields. '
        f'{summary["Stock Quantity Missing"]:,} records still need a stock quantity.</div>',
        unsafe_allow_html=True,
    )

    chart_columns = st.columns(3)
    with chart_columns[0]:
        st.subheader("Stock status")
        status_counts = inventory["stock_status"].value_counts().rename("Records")
        st.bar_chart(status_counts, color="#00A6A6", height=300)
    with chart_columns[1]:
        st.subheader("Records by brand")
        brands = (
            inventory.assign(brand=inventory["brand"].fillna("Not Set"))
            .groupby("brand", dropna=False)
            .size()
            .sort_values(ascending=False)
            .head(10)
            .rename("Records")
        )
        st.bar_chart(brands, color="#173F5F", height=300)
    with chart_columns[2]:
        st.subheader("Upcoming expiries")
        expiry = inventory.loc[
            inventory["expiry_date"].notna() & inventory["days_to_expiry"].between(0, 365)
        ].copy()
        if expiry.empty:
            st.info("No dated expiries in the next 12 months.")
        else:
            expiry["Expiry Month"] = expiry["expiry_date"].dt.to_period("M").astype(str)
            expiry_chart = expiry.groupby("Expiry Month").size().rename("Records")
            st.bar_chart(expiry_chart, color="#F4B942", height=300)

    st.subheader("Priority attention")
    priority_order = {
        "Expired": 0,
        "Expiring Soon": 1,
        "Out of Stock": 2,
        "Low Stock": 3,
        "Stock Not Set": 4,
    }
    priority = inventory.loc[
        inventory["overall_status"].isin(priority_order)
    ].copy()
    priority["_priority"] = priority["overall_status"].map(priority_order)
    priority = priority.sort_values(
        ["_priority", "expiry_date", "item_name"], na_position="last"
    ).head(20)
    priority_columns = [
        "inventory_code",
        "item_name",
        "brand",
        "quantity_in_stock",
        "reorder_level",
        "expiry_date",
        "days_to_expiry",
        "overall_status",
    ]
    st.dataframe(
        display_frame(priority[priority_columns]),
        use_container_width=True,
        hide_index=True,
        height=430,
    )


def render_inventory(inventory: pd.DataFrame, warning_days: int) -> None:
    page_header(
        "Inventory List",
        "Search the source records, update key fields in bulk, or maintain one record in detail.",
    )
    browse_tab, add_tab, edit_tab = st.tabs(["Browse & Quick Update", "Add New Item", "Edit / Archive"])

    with browse_tab:
        include_archived = st.toggle("Include archived records", value=False)
        frame = fetch_inventory_dataframe(
            include_archived=include_archived, warning_days=warning_days
        )
        filtered = apply_inventory_filters(frame, "inventory")
        st.caption(f"Showing {len(filtered):,} of {len(frame):,} records")
        browse_columns = [
            "inventory_code",
            "item_name",
            "brand",
            "description",
            "unit_price",
            "quantity_in_stock",
            "reorder_level",
            "expiry_date",
            "stock_status",
            "expiry_status",
            "inventory_value",
        ]
        st.dataframe(
            display_frame(filtered[browse_columns]),
            use_container_width=True,
            hide_index=True,
            height=470,
        )

        with st.expander("Quick-update stock, reorder, price and expiry fields"):
            if filtered.empty:
                st.info("No records match the current filters.")
            else:
                editable = filtered.head(250).loc[
                    :,
                    [
                        "id",
                        "inventory_code",
                        "item_name",
                        "brand",
                        "quantity_in_stock",
                        "reorder_level",
                        "reorder_quantity",
                        "unit_price",
                        "selling_price_per_test",
                        "expiry_date",
                    ],
                ].copy()
                editable["expiry_date"] = pd.to_datetime(
                    editable["expiry_date"], errors="coerce"
                ).dt.date
                if len(filtered) > len(editable):
                    st.caption("Quick editor is limited to the first 250 filtered records.")
                edited = st.data_editor(
                    editable,
                    hide_index=True,
                    use_container_width=True,
                    num_rows="fixed",
                    disabled=["id", "inventory_code", "item_name", "brand"],
                    column_config={
                        "id": st.column_config.NumberColumn("ID"),
                        "inventory_code": st.column_config.TextColumn("Inventory Code"),
                        "item_name": st.column_config.TextColumn("Item Name"),
                        "brand": st.column_config.TextColumn("Brand"),
                        "quantity_in_stock": st.column_config.NumberColumn(
                            "Quantity In Stock", min_value=0.0
                        ),
                        "reorder_level": st.column_config.NumberColumn(
                            "Reorder Level", min_value=0.0
                        ),
                        "reorder_quantity": st.column_config.NumberColumn(
                            "Reorder Quantity", min_value=0.0
                        ),
                        "unit_price": st.column_config.NumberColumn(
                            "Unit Price", min_value=0.0, format="%.2f"
                        ),
                        "selling_price_per_test": st.column_config.NumberColumn(
                            "Selling Price / Test", min_value=0.0, format="%.2f"
                        ),
                        "expiry_date": st.column_config.DateColumn("Expiry Date"),
                    },
                    key="quick_inventory_editor",
                )
                if st.button("Save quick updates", type="primary", key="save_quick"):
                    try:
                        rows = edited.to_dict(orient="records")
                        updated = bulk_update_items(rows)
                        st.success(f"Saved {updated:,} inventory records.")
                    except ValueError as exc:
                        st.error(str(exc))

    with add_tab:
        st.subheader("Create inventory record")
        with st.form("add_inventory_form", clear_on_submit=True):
            first = st.columns(4)
            code = first[0].text_input("Inventory code", placeholder="Auto-generated if blank")
            item_name = first[1].text_input("Item name *")
            brand = first[2].text_input("Brand")
            category = first[3].text_input("Category", value="Laboratory Inventory")

            second = st.columns(4)
            test_volume = second[0].text_input("Test volume")
            description = second[1].text_input("Description / pack size")
            batch_no = second[2].text_input("Batch / lot no.")
            location = second[3].text_input("Storage location")

            third = st.columns(4)
            unit_price = third[0].text_input("Unit price", placeholder="Blank = unknown")
            total_tests = third[1].text_input("Total test effective", placeholder="Blank = unknown")
            selling_price = third[2].text_input("Selling price per test", placeholder="Blank = unknown")
            quantity = third[3].text_input("Quantity in stock", placeholder="Blank = not set")

            fourth = st.columns(4)
            reorder_level = fourth[0].text_input("Reorder level", placeholder="Blank = not set")
            reorder_quantity = fourth[1].text_input("Reorder quantity", placeholder="Blank = not set")
            supplier = fourth[2].text_input("Supplier")
            reminder_days = fourth[3].text_input("Reminder days", placeholder="Blank = not set")

            date_columns = st.columns(3)
            has_opened = date_columns[0].checkbox("Opened date available", value=False)
            opened_date = date_columns[0].date_input(
                "Opened date", value=date.today(), disabled=not has_opened
            )
            has_expiry = date_columns[1].checkbox("Expiry date available", value=True)
            expiry_date = date_columns[1].date_input(
                "Expiry date", value=date.today(), disabled=not has_expiry
            )
            has_reorder_due = date_columns[2].checkbox("Reorder due date available", value=False)
            reorder_due = date_columns[2].date_input(
                "Reorder due date", value=date.today(), disabled=not has_reorder_due
            )
            notes = st.text_area("Notes")
            submitted = st.form_submit_button("Add inventory record", type="primary")
        if submitted:
            try:
                item_id = add_item(
                    {
                        "inventory_code": code,
                        "item_name": item_name,
                        "brand": brand,
                        "category": category,
                        "test_volume": test_volume,
                        "description": description,
                        "batch_no": batch_no,
                        "storage_location": location,
                        "unit_price": optional_number(unit_price, "Unit price"),
                        "total_test_effective": optional_number(total_tests, "Total test effective"),
                        "selling_price_per_test": optional_number(selling_price, "Selling price per test"),
                        "quantity_in_stock": optional_number(quantity, "Quantity in stock"),
                        "reorder_level": optional_number(reorder_level, "Reorder level"),
                        "reorder_quantity": optional_number(reorder_quantity, "Reorder quantity"),
                        "supplier": supplier,
                        "reminder_days": optional_number(reminder_days, "Reminder days"),
                        "opened_date": opened_date if has_opened else None,
                        "expiry_date": expiry_date if has_expiry else None,
                        "reorder_due_date": reorder_due if has_reorder_due else None,
                        "notes": notes,
                        "is_active": 1,
                    }
                )
                st.success(f"Inventory record created (database ID {item_id}).")
            except ValueError as exc:
                st.error(str(exc))

    with edit_tab:
        if inventory.empty:
            st.info("No records are available to edit.")
        else:
            all_records = fetch_inventory_dataframe(include_archived=True, warning_days=warning_days)
            option_labels = {
                int(row.id): f"{row.inventory_code} · {row.item_name} · {text_value(row.brand) or 'No brand'}"
                for row in all_records.itertuples()
            }
            selected_id = st.selectbox(
                "Select record",
                options=list(option_labels),
                format_func=lambda item_id: option_labels[item_id],
            )
            selected = all_records.loc[all_records["id"].eq(selected_id)].iloc[0]
            suffix = str(selected_id)
            with st.form(f"edit_inventory_{suffix}"):
                first = st.columns(4)
                edit_code = first[0].text_input(
                    "Inventory code", value=str(selected["inventory_code"]), key=f"edit_code_{suffix}"
                )
                edit_name = first[1].text_input(
                    "Item name *", value=str(selected["item_name"]), key=f"edit_name_{suffix}"
                )
                edit_brand = first[2].text_input(
                    "Brand", value=text_value(selected["brand"]), key=f"edit_brand_{suffix}"
                )
                edit_category = first[3].text_input(
                    "Category", value=text_value(selected["category"]), key=f"edit_category_{suffix}"
                )
                second = st.columns(4)
                edit_test_volume = second[0].text_input(
                    "Test volume", value=text_value(selected["test_volume"]), key=f"edit_volume_{suffix}"
                )
                edit_description = second[1].text_input(
                    "Description", value=text_value(selected["description"]), key=f"edit_desc_{suffix}"
                )
                edit_batch = second[2].text_input(
                    "Batch / lot no.", value=text_value(selected["batch_no"]), key=f"edit_batch_{suffix}"
                )
                edit_location = second[3].text_input(
                    "Storage location", value=text_value(selected["storage_location"]), key=f"edit_loc_{suffix}"
                )
                third = st.columns(4)
                edit_unit_price = third[0].text_input(
                    "Unit price", value=numeric_text(selected["unit_price"]), key=f"edit_price_{suffix}"
                )
                edit_total_tests = third[1].text_input(
                    "Total test effective",
                    value=numeric_text(selected["total_test_effective"]),
                    key=f"edit_tests_{suffix}",
                )
                edit_selling = third[2].text_input(
                    "Selling price per test",
                    value=numeric_text(selected["selling_price_per_test"]),
                    key=f"edit_sell_{suffix}",
                )
                edit_quantity = third[3].text_input(
                    "Quantity in stock",
                    value=numeric_text(selected["quantity_in_stock"]),
                    key=f"edit_qty_{suffix}",
                )
                fourth = st.columns(4)
                edit_reorder_level = fourth[0].text_input(
                    "Reorder level",
                    value=numeric_text(selected["reorder_level"]),
                    key=f"edit_reorder_{suffix}",
                )
                edit_reorder_qty = fourth[1].text_input(
                    "Reorder quantity",
                    value=numeric_text(selected["reorder_quantity"]),
                    key=f"edit_reorder_qty_{suffix}",
                )
                edit_supplier = fourth[2].text_input(
                    "Supplier", value=text_value(selected["supplier"]), key=f"edit_supplier_{suffix}"
                )
                edit_reminder = fourth[3].text_input(
                    "Reminder days",
                    value=numeric_text(selected["reminder_days"]),
                    key=f"edit_reminder_{suffix}",
                )
                dates = st.columns(3)
                opened_known = not pd.isna(selected["opened_date"])
                edit_has_opened = dates[0].checkbox(
                    "Opened date available", value=opened_known, key=f"edit_has_opened_{suffix}"
                )
                edit_opened = dates[0].date_input(
                    "Opened date",
                    value=date_value(selected["opened_date"]),
                    disabled=not edit_has_opened,
                    key=f"edit_opened_{suffix}",
                )
                expiry_known = not pd.isna(selected["expiry_date"])
                edit_has_expiry = dates[1].checkbox(
                    "Expiry date available", value=expiry_known, key=f"edit_has_expiry_{suffix}"
                )
                edit_expiry = dates[1].date_input(
                    "Expiry date",
                    value=date_value(selected["expiry_date"]),
                    disabled=not edit_has_expiry,
                    key=f"edit_expiry_{suffix}",
                )
                due_known = not pd.isna(selected["reorder_due_date"])
                edit_has_due = dates[2].checkbox(
                    "Reorder due date available", value=due_known, key=f"edit_has_due_{suffix}"
                )
                edit_due = dates[2].date_input(
                    "Reorder due date",
                    value=date_value(selected["reorder_due_date"]),
                    disabled=not edit_has_due,
                    key=f"edit_due_{suffix}",
                )
                edit_notes = st.text_area(
                    "Notes", value=text_value(selected["notes"]), key=f"edit_notes_{suffix}"
                )
                edit_submitted = st.form_submit_button("Save record", type="primary")
            if edit_submitted:
                try:
                    update_item(
                        selected_id,
                        {
                            "inventory_code": edit_code,
                            "item_name": edit_name,
                            "brand": edit_brand,
                            "category": edit_category,
                            "test_volume": edit_test_volume,
                            "description": edit_description,
                            "batch_no": edit_batch,
                            "storage_location": edit_location,
                            "unit_price": optional_number(edit_unit_price, "Unit price"),
                            "total_test_effective": optional_number(edit_total_tests, "Total test effective"),
                            "selling_price_per_test": optional_number(edit_selling, "Selling price per test"),
                            "quantity_in_stock": optional_number(edit_quantity, "Quantity in stock"),
                            "reorder_level": optional_number(edit_reorder_level, "Reorder level"),
                            "reorder_quantity": optional_number(edit_reorder_qty, "Reorder quantity"),
                            "supplier": edit_supplier,
                            "reminder_days": optional_number(edit_reminder, "Reminder days"),
                            "opened_date": edit_opened if edit_has_opened else None,
                            "expiry_date": edit_expiry if edit_has_expiry else None,
                            "reorder_due_date": edit_due if edit_has_due else None,
                            "notes": edit_notes,
                        },
                    )
                    st.success("Inventory record updated.")
                except ValueError as exc:
                    st.error(str(exc))

            is_active = int(selected["is_active"]) == 1
            action = "Archive" if is_active else "Restore"
            st.divider()
            confirmed = st.checkbox(
                f"I confirm that I want to {action.lower()} this record",
                key=f"archive_confirm_{suffix}",
            )
            if st.button(
                f"{action} record",
                disabled=not confirmed,
                type="secondary",
                key=f"archive_button_{suffix}",
            ):
                archive_item(selected_id, archived=is_active)
                st.success(f"Record {action.lower()}d.")


def render_stock_update(inventory: pd.DataFrame, warning_days: int) -> None:
    page_header(
        "Stock Update",
        "Receive, issue or adjust stock with a permanent movement history.",
    )
    if inventory.empty:
        st.info("Add an active inventory record before recording stock movements.")
        return
    labels = {
        int(row.id): f"{row.inventory_code} · {row.item_name} · {text_value(row.brand) or 'No brand'}"
        for row in inventory.itertuples()
    }
    item_id = st.selectbox(
        "Inventory record",
        options=list(labels),
        format_func=lambda selected_id: labels[selected_id],
    )
    selected = inventory.loc[inventory["id"].eq(item_id)].iloc[0]
    metric_columns = st.columns(4)
    metric_columns[0].metric("Current Quantity", numeric_text(selected["quantity_in_stock"]) or "Not set")
    metric_columns[1].metric("Reorder Level", numeric_text(selected["reorder_level"]) or "Not set")
    metric_columns[2].metric("Stock Status", selected["stock_status"])
    metric_columns[3].metric("Expiry Status", selected["expiry_status"])

    with st.form("stock_movement_form", clear_on_submit=True):
        columns = st.columns(3)
        movement_type = columns[0].selectbox("Movement type", ["Receive", "Issue", "Adjustment"])
        quantity_label = "New stock balance" if movement_type == "Adjustment" else "Quantity"
        quantity = columns[1].number_input(quantity_label, min_value=0.0, value=0.0, step=1.0)
        movement_date = columns[2].date_input("Movement date", value=date.today())
        details = st.columns(2)
        reference = details[0].text_input("Reference", placeholder="PO, invoice, request or adjustment reference")
        notes = details[1].text_input("Notes")
        movement_submit = st.form_submit_button("Save stock movement", type="primary")
    if movement_submit:
        try:
            new_balance = record_stock_movement(
                item_id,
                movement_type,
                quantity,
                reference=reference,
                notes=notes,
                moved_at=movement_date,
            )
            st.success(f"Stock movement saved. New balance: {new_balance:g}.")
        except ValueError as exc:
            st.error(str(exc))

    st.subheader("Recent movement history")
    movements = fetch_movements_dataframe()
    movements = movements.loc[movements["inventory_code"].eq(selected["inventory_code"])].head(100)
    if movements.empty:
        st.info("No stock movements have been recorded for this item.")
    else:
        st.dataframe(
            display_frame(
                movements[
                    [
                        "moved_at",
                        "movement_type",
                        "quantity_change",
                        "quantity_before",
                        "quantity_after",
                        "reference",
                        "notes",
                    ]
                ]
            ),
            use_container_width=True,
            hide_index=True,
        )


def render_reports(inventory: pd.DataFrame, currency: str) -> None:
    page_header(
        "Reports",
        "View filtered reports and download formatted Excel or CSV files.",
    )
    report_name = st.selectbox(
        "Report type",
        [
            "Full Inventory Report",
            "Reorder Report",
            "Expiry Report",
            "Inventory Valuation Report",
            "Data Quality Report",
            "Stock Movement Report",
        ],
    )
    movements = fetch_movements_dataframe()
    filtered_inventory = inventory.copy()
    filter_columns = st.columns([2, 1.2, 1.2])
    search = filter_columns[0].text_input("Search report", placeholder="Item, code or brand")
    brands = sorted(filtered_inventory["brand"].dropna().astype(str).unique())
    report_brands = filter_columns[1].multiselect("Brand", brands)
    categories = sorted(filtered_inventory["category"].dropna().astype(str).unique())
    report_categories = filter_columns[2].multiselect("Category", categories)
    if search.strip() and not filtered_inventory.empty:
        pattern = re.escape(search.strip())
        combined = filtered_inventory[["inventory_code", "item_name", "brand"]].fillna("").astype(str).agg(" ".join, axis=1)
        filtered_inventory = filtered_inventory.loc[combined.str.contains(pattern, case=False, regex=True)]
        if not movements.empty:
            movement_search = movements[["inventory_code", "item_name", "brand"]].fillna("").astype(str).agg(" ".join, axis=1)
            movements = movements.loc[movement_search.str.contains(pattern, case=False, regex=True)]
    if report_brands:
        filtered_inventory = filtered_inventory.loc[filtered_inventory["brand"].isin(report_brands)]
        if not movements.empty:
            movements = movements.loc[movements["brand"].isin(report_brands)]
    if report_categories:
        filtered_inventory = filtered_inventory.loc[filtered_inventory["category"].isin(report_categories)]

    report_frame = build_report_frame(report_name, filtered_inventory, movements)
    summary = inventory_summary(filtered_inventory)
    st.caption(f"{len(report_frame):,} report rows")
    st.dataframe(
        display_frame(report_frame),
        use_container_width=True,
        hide_index=True,
        height=510,
    )
    excel = create_excel_report(report_name, report_frame, summary, currency)
    slug = re.sub(r"[^a-z0-9]+", "_", report_name.lower()).strip("_")
    download_columns = st.columns(2)
    download_columns[0].download_button(
        "Download Excel report",
        data=excel,
        file_name=f"{slug}_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    download_columns[1].download_button(
        "Download CSV report",
        data=csv_bytes(report_frame),
        file_name=f"{slug}_{date.today().isoformat()}.csv",
        mime="text/csv",
        use_container_width=True,
    )


def render_import_backup(inventory: pd.DataFrame, currency: str) -> None:
    page_header(
        "Import & Backup",
        "Import updated lists, download a clean template, and keep recoverable backups.",
    )
    st.subheader("Import inventory")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx", "xlsm", "csv"])
    if uploaded is not None:
        try:
            records, diagnostics = parse_tabular_upload(uploaded.name, uploaded.getvalue())
            metric_columns = st.columns(3)
            metric_columns[0].metric("Rows Ready", f"{diagnostics['records']:,}")
            metric_columns[1].metric("Missing Stock", f"{diagnostics['records_missing_stock']:,}")
            metric_columns[2].metric("Missing Expiry", f"{diagnostics['records_missing_expiry']:,}")
            st.caption("Mapped fields: " + ", ".join(diagnostics["mapped_fields"]))
            st.dataframe(records_preview(records), use_container_width=True, hide_index=True)
            mode = st.radio("Import mode", ["Append to current inventory", "Replace current inventory"])
            confirm_replace = True
            if mode == "Replace current inventory":
                st.warning("Replace mode removes the current inventory and its stock-movement history.")
                confirm_replace = st.checkbox("I have downloaded a backup and confirm replacement")
            if st.button(
                "Import records",
                type="primary",
                disabled=not records or not confirm_replace,
            ):
                result = insert_items(records, replace=mode == "Replace current inventory")
                st.success(
                    f"Import complete: {result['inserted']:,} inserted, {result['skipped']:,} skipped."
                )
        except (ValueError, OSError) as exc:
            st.error(str(exc))

    st.divider()
    st.subheader("Downloads and backup")
    st.caption("Use the template for future imports. Keep database backups before bulk replacements.")
    full_report = build_report_frame("Full Inventory Report", inventory)
    summary = inventory_summary(inventory)
    full_excel = create_excel_report("Full Inventory Report", full_report, summary, currency)
    backup_columns = st.columns(3)
    backup_columns[0].download_button(
        "Download import template",
        data=create_import_template(),
        file_name="inventory_import_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    backup_columns[1].download_button(
        "Download all inventory",
        data=full_excel,
        file_name=f"inventory_export_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    backup_columns[2].download_button(
        "Download database backup",
        data=create_database_backup(),
        file_name=f"inventory_backup_{date.today().isoformat()}.db",
        mime="application/octet-stream",
        use_container_width=True,
    )


def render_settings(inventory: pd.DataFrame, settings: dict[str, str]) -> None:
    page_header(
        "Settings",
        "Configure report identity, currency display and expiry alert timing.",
    )
    with st.form("settings_form"):
        columns = st.columns(3)
        business_name = columns[0].text_input(
            "Business / facility name", value=settings.get("business_name", "Inventory Control Centre")
        )
        currency_symbol = columns[1].text_input(
            "Currency symbol", value=settings.get("currency_symbol", "₹"), max_chars=5
        )
        expiry_days = columns[2].number_input(
            "Expiring-soon warning (days)",
            min_value=1,
            max_value=365,
            value=int(settings.get("expiry_warning_days", "30")),
        )
        save = st.form_submit_button("Save settings", type="primary")
    if save:
        save_settings(
            {
                "business_name": business_name.strip() or "Inventory Control Centre",
                "currency_symbol": currency_symbol.strip() or "₹",
                "expiry_warning_days": int(expiry_days),
            }
        )
        st.success("Settings saved. Refresh the page to apply the new heading and alert window.")

    st.subheader("Database health")
    summary = inventory_summary(inventory)
    health = pd.DataFrame(
        {
            "Check": [
                "Active inventory records",
                "Records with stock quantity",
                "Records with expiry date",
                "Records with unit price",
                "Records needing source-value review",
            ],
            "Count": [
                len(inventory),
                int(inventory["quantity_in_stock"].notna().sum()) if not inventory.empty else 0,
                int(inventory["expiry_date"].notna().sum()) if not inventory.empty else 0,
                int(inventory["unit_price"].notna().sum()) if not inventory.empty else 0,
                int(inventory["source_notes"].fillna("").astype(str).str.len().gt(0).sum())
                if not inventory.empty
                else 0,
            ],
        }
    )
    st.dataframe(health, use_container_width=True, hide_index=True)
    st.caption(
        f"Current known inventory value: {money(summary['Known Inventory Value'], settings.get('currency_symbol', '₹'))}."
    )


def main() -> None:
    initialize_app_data()
    settings = get_settings()
    warning_days = int(settings.get("expiry_warning_days", "30"))
    currency = settings.get("currency_symbol", "₹")
    inventory = fetch_inventory_dataframe(warning_days=warning_days)

    st.sidebar.markdown("## 📦 Inventory")
    st.sidebar.caption(settings.get("business_name", "Inventory Control Centre"))
    page = st.sidebar.radio(
        "Navigation",
        [
            "Dashboard",
            "Inventory List",
            "Stock Update",
            "Reports",
            "Import & Backup",
            "Settings",
        ],
        label_visibility="collapsed",
    )
    st.sidebar.divider()
    st.sidebar.caption(f"{len(inventory):,} active records · SQLite database")

    seed_result = st.session_state.pop("seed_result", None)
    if seed_result:
        st.success(
            f"Initial workbook imported: {seed_result['inserted']:,} records. "
            f"{seed_result['records_missing_stock']:,} records are flagged for stock entry."
        )

    if page == "Dashboard":
        render_dashboard(inventory, currency)
    elif page == "Inventory List":
        render_inventory(inventory, warning_days)
    elif page == "Stock Update":
        render_stock_update(inventory, warning_days)
    elif page == "Reports":
        render_reports(inventory, currency)
    elif page == "Import & Backup":
        render_import_backup(inventory, currency)
    else:
        render_settings(inventory, settings)


if __name__ == "__main__":
    main()
