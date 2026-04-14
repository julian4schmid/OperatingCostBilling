import pandas as pd
import psycopg2
from config import DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD, DATA_DIR


def get_connection():
    return psycopg2.connect(
        host=DB_HOST,
        port=DB_PORT,
        dbname=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD
    )


def import_costs(year: int, overwrite: bool = False):
    file_path = DATA_DIR / f"Kosten{year}.xlsx"

    conn = get_connection()
    cur = conn.cursor()

    # =========================
    # Prefetch existing data
    # =========================
    cur.execute("SELECT building_id, year, cost_type, amount FROM costs WHERE year = %s", (year,))
    existing_costs = {(r[0], r[1], r[2]): r[3] for r in cur.fetchall()}

    cur.execute("""
        SELECT building_id, year, unit_id, cost_type, amount, allocation_key
        FROM individual_costs
        WHERE year = %s
    """, (year,))
    existing_individual = {(r[0], r[1], r[2], r[3]): (r[4], r[5]) for r in cur.fetchall()}

    cur.execute("SELECT building_id, cost_type FROM building_cost_allocation")
    allocation_set = set(cur.fetchall())

    # Track what appears in Excel
    seen_building_keys = set()
    seen_individual_keys = set()

    # Batch containers
    building_cost_rows = []
    building_cost_deletes = []

    individual_rows = []
    individual_deletes = []

    # =========================
    # PROCESS "Allgemein"
    # =========================
    df_gen = pd.read_excel(file_path, sheet_name="Allgemein")

    cost_types = df_gen.iloc[1:, 0]
    building_ids = df_gen.columns[1:]

    for row_idx, cost_type in enumerate(cost_types, start=1):
        cost_type = str(cost_type).strip()

        for col_idx, building_id in enumerate(building_ids, start=1):
            building_id = str(building_id).strip()
            amount = df_gen.iloc[row_idx, col_idx]

            key = (building_id, year, cost_type)
            seen_building_keys.add(key)

            existing_value = existing_costs.get(key)

            # --- NaN → delete ---
            if pd.isna(amount):
                if existing_value is not None:
                    print(f"⚠️ DELETE: {cost_type} | {building_id} | {year} → DB: {existing_value}")
                    building_cost_deletes.append((building_id, year, cost_type))
                continue

            amount = float(amount)

            # --- Insert if not existing ---
            if existing_value is None:
                building_cost_rows.append((building_id, year, cost_type, amount))

                if (building_id, cost_type) not in allocation_set:
                    print(f"⚠️ {cost_type} ({building_id}, {year}) has no allocation type stored")

                continue

            # --- Same value → skip ---
            if existing_value == amount:
                continue

            # --- Different value ---
            print(f"⚠️ CHANGE: {cost_type} | {building_id} | {year} → {existing_value} → {amount}")

            if overwrite:
                building_cost_rows.append((building_id, year, cost_type, amount))
            else:
                continue

            if (building_id, cost_type) not in allocation_set:
                print(f"⚠️ {cost_type} ({building_id}, {year}) has no allocation type stored")

    # =========================
    # PROCESS "Individual"
    # =========================
    df_ind = pd.read_excel(file_path, sheet_name="Individuell")

    for _, row in df_ind.iloc[1:].iterrows():
        cost_id = str(row[0]).strip()
        building_id = str(row[1]).strip()
        unit_id = str(row[2]).strip()
        cost_type = str(row[3]).strip()
        allocation_key = str(row[4]).strip()
        amount = row[5]
        date_val = row[6]
        usage = row[7]
        price = row[8]

        # --- normalize date ---
        if pd.isna(date_val):
            date_val = None
        else:
            date_val = pd.to_datetime(date_val).date()

        # --- normalize numeric fields ---
        usage = None if pd.isna(usage) else float(usage)
        price = None if pd.isna(price) else float(price)

        seen_individual_keys.add(cost_id)

        existing = existing_individual.get(cost_id)
        if existing:
            existing_amount, existing_alloc_key = existing
        else:
            existing_amount, existing_alloc_key = (None, None)

        # =========================
        # DELETE (amount is NaN)
        # =========================
        if pd.isna(amount):
            if existing_amount is not None:
                print(f"⚠️ DELETE: {cost_type} | {building_id} | {unit_id} | {year} → DB: {existing_amount}")
                individual_deletes.append((cost_id,))
            continue

        amount = float(amount)

        # =========================
        # INSERT
        # =========================
        if existing_amount is None:
            individual_rows.append((
                cost_id,
                building_id,
                year,
                unit_id,
                cost_type,
                amount,
                allocation_key,
                date_val,
                usage,
                price
            ))
            continue

        # =========================
        # SAME VALUE
        # =========================
        if existing_amount == amount:
            if existing_alloc_key != allocation_key:
                if overwrite:
                    print(f"⚠️ UPDATE allocation_key: {cost_type} | {building_id} | {unit_id} | {year} → "
                          f"{existing_alloc_key} → {allocation_key}")
                    individual_rows.append((
                        cost_id,
                        building_id,
                        year,
                        unit_id,
                        cost_type,
                        amount,
                        allocation_key,
                        date_val,
                        usage,
                        price
                    ))
                else:
                    print(
                        f"⚠️ allocation_key differs but overwrite=False: {cost_type} | {building_id} | {unit_id} | {year} → "
                        f"{existing_alloc_key} → {allocation_key}")
            continue

        # =========================
        # DIFFERENT VALUE
        # =========================
        print(f"⚠️ CHANGE: {cost_type} | {building_id} | {unit_id} | {year} → {existing_amount} → {amount}")

        if overwrite:
            individual_rows.append((
                cost_id,
                building_id,
                year,
                unit_id,
                cost_type,
                amount,
                allocation_key,
                date_val,
                usage,
                price
            ))
        else:
            continue

    # =========================
    # WARN: DB entries missing in Excel
    # =========================
    for key, value in existing_costs.items():
        if key not in seen_building_keys:
            print(f"⚠️ MISSING IN EXCEL: {key} → DB: {value}")

    for key, value in existing_individual.items():
        if key not in seen_individual_keys:
            print(f"⚠️ MISSING IN EXCEL: {key} → DB: {value}")

    # =========================
    # EXECUTE DELETES
    # =========================
    if building_cost_deletes:
        cur.executemany("""
            DELETE FROM costs
            WHERE building_id = %s
              AND year = %s
              AND cost_type = %s
        """, building_cost_deletes)

    if individual_deletes:
        cur.executemany("""
            DELETE FROM individual_costs
            WHERE cost_id = %s
        """, individual_deletes)

    # =========================
    # EXECUTE UPSERTS
    # =========================
    if building_cost_rows:
        cur.executemany("""
            INSERT INTO costs (building_id, year, cost_type, amount)
            VALUES (%s, %s, %s, %s)
            ON CONFLICT (building_id, year, cost_type)
            DO UPDATE SET amount = EXCLUDED.amount
        """, building_cost_rows)

    if individual_rows:
        cur.executemany("""
            INSERT INTO individual_costs (
                cost_id,
                building_id,
                year,
                unit_id,
                cost_type,
                amount,
                allocation_key,
                date,
                usage,
                price
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (cost_id)
            DO UPDATE SET
                building_id = EXCLUDED.building_id,
                year = EXCLUDED.year,
                unit_id = EXCLUDED.unit_id,
                cost_type = EXCLUDED.cost_type,
                amount = EXCLUDED.amount,
                allocation_key = EXCLUDED.allocation_key,
                date = EXCLUDED.date,
                usage = EXCLUDED.usage,
                price = EXCLUDED.price
        """, individual_rows)

    conn.commit()
    cur.close()
    conn.close()

    print(f"✅ Import for {year} completed")