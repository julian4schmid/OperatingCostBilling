import pandas as pd
from psycopg2.extras import execute_values
from yearly_import import get_connection
from config import DATA_DIR

EXCEL_FILE = DATA_DIR / "data.xlsx"

TABLES = [
    "buildings",
    "units",
    "tenants",
    "building_cost_allocations"
]


def clean_dataframe(df):
    # Remove completely empty rows
    df = df.dropna(how="all")

    # Convert NaN to None for SQL
    df = df.where(pd.notnull(df), None)

    return df


def upsert_table(conn, table_name, df, overwrite=False):
    if df.empty:
        print(f"⚠️ {table_name}: no data")
        return

    columns = list(df.columns)
    values = [tuple(row) for row in df.to_numpy()]

    cols_str = ", ".join(columns)

    # Define primary keys
    if table_name == "buildings":
        conflict_cols = ["building_id"]
    elif table_name == "units":
        conflict_cols = ["building_id", "unit_position"]
    elif table_name == "tenants":
        conflict_cols = ["tenant_id"]
    elif table_name == "building_cost_allocations":
        conflict_cols = ["allocation_id"]
    else:
        raise ValueError(f"Unknown table: {table_name}")

    conflict_str = ", ".join(conflict_cols)

    # Columns to update (exclude primary keys)
    update_cols = [col for col in columns if col not in conflict_cols]

    if overwrite:
        update_str = ", ".join([f"{col} = EXCLUDED.{col}" for col in update_cols])

        query = f"""
            INSERT INTO {table_name} ({cols_str})
            VALUES %s
            ON CONFLICT ({conflict_str}) DO UPDATE
            SET {update_str};
        """
    else:
        # Do nothing on conflict, only warn later
        query = f"""
            INSERT INTO {table_name} ({cols_str})
            VALUES %s
            ON CONFLICT ({conflict_str}) DO NOTHING;
        """

    with conn.cursor() as cur:
        execute_values(cur, query, values)

        if not overwrite:
            # Check how many rows were actually inserted
            inserted = cur.rowcount
            total = len(values)
            skipped = total - inserted

            if skipped > 0:
                print(f"⚠️ {table_name}: {skipped} duplicate rows skipped (overwrite=False)")

    print(f"✅ {table_name}: {len(values)} rows processed")


def import_table(conn, table: str, overwrite=False):
    print(f"\n📥 Import {table}...")

    df = pd.read_excel(EXCEL_FILE, sheet_name=table)
    df = clean_dataframe(df)

    upsert_table(conn, table, df, overwrite=overwrite)


def import_all(overwrite=False):
    conn = get_connection()

    try:
        for table in TABLES:
            import_table(conn, table, overwrite=overwrite)

        conn.commit()
        print("\n🎉 All imports completed")

    except Exception as e:
        conn.rollback()
        print("❌ Error:", e)

    finally:
        conn.close()

