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
        SELECT building_id, year, unit_position, cost_type, amount, allocation_key
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
    # PROCESS "Individuell"
    # =========================
    df_ind = pd.read_excel(file_path, sheet_name="Individuell")

    for _, row in df_ind.iloc[1:].iterrows():
        building_id = str(row[0]).strip()
        unit_position = str(row[1]).strip()
        cost_type = str(row[2]).strip()
        allocation_key = str(row[3]).strip()
        amount = row[4]

        key = (building_id, year, unit_position, cost_type)
        seen_individual_keys.add(key)

        existing_amount, existing_alloc_key = existing_individual.get(key, (None, None))


        # --- NaN → delete ---
        if pd.isna(amount):
            if existing_amount is not None:
                print(f"⚠️ DELETE: {cost_type} | {building_id} | {unit_position} | {year} → DB: {existing_amount}")
                individual_deletes.append((building_id, year, unit_position, cost_type))
            continue

        amount = float(amount)

        # --- Insert if not existing ---
        if existing_amount is None:
            individual_rows.append((building_id, year, unit_position, cost_type, amount, allocation_key))
            continue

        # --- Same value / allocation_key check ---
        if existing_amount == amount:
            if existing_alloc_key != allocation_key:
                # Allocation key changed → update only if overwrite=True
                if overwrite:
                    print(f"⚠️ UPDATE allocation_key: {cost_type} | {building_id} | {unit_position} | {year} → "
                          f"{existing_alloc_key} → {allocation_key}")
                    individual_rows.append((building_id, year, unit_position, cost_type, amount, allocation_key))
                else:
                    print(
                        f"⚠️ allocation_key differs but overwrite=False: {cost_type} | {building_id} | {unit_position} | {year} → "
                        f"{existing_alloc_key} → {allocation_key}")
            # Amount same → skip further processing
            continue

        # --- Different value ---
        print(f"⚠️ CHANGE: {cost_type} | {building_id} | {unit_position} | {year} → {existing_amount} → {amount}")

        if overwrite:
            individual_rows.append((building_id, year, unit_position, cost_type, amount, allocation_key))
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
            WHERE building_id = %s
              AND year = %s
              AND unit_position = %s
              AND cost_type = %s
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
            INSERT INTO individual_costs (building_id, year, unit_position, cost_type, amount, allocation_key)
            VALUES (%s, %s, %s, %s, %s, %s)
            ON CONFLICT (building_id, year, unit_position, cost_type)
            DO UPDATE SET 
                amount = EXCLUDED.amount,
                allocation_key = EXCLUDED.allocation_key
        """, individual_rows)

    conn.commit()
    cur.close()
    conn.close()

    print(f"✅ Import for {year} completed")