from datetime import date, timedelta
from psycopg2.extras import RealDictCursor
from yearly_import import get_connection



# =========================
# MAIN ENTRY
# =========================

def calculate_billing(building_id: str, year: int):
    conn = get_connection()

    try:
        data = load_data(conn, building_id, year)
        allocation_map = build_allocation_map(data)

        results = []

        for tenant in data["tenants"]:
            result = calculate_for_tenant(
                tenant,
                data,
                allocation_map,
                year
            )
            results.append(result)

        return results

    finally:
        conn.close()

# =========================
# DATA LOADING
# =========================

def load_data(conn, building_id: str, year: int):
    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        # Building
        cur.execute("""
            SELECT *
            FROM buildings
            WHERE building_id = %s
        """, (building_id,))
        building = cur.fetchone()

        # Units
        cur.execute("""
            SELECT *
            FROM units
            WHERE building_id = %s
        """, (building_id,))
        units = cur.fetchall()

        # Tenants
        cur.execute("""
            SELECT *
            FROM tenants
            WHERE building_id = %s
        """, (building_id,))
        tenants = cur.fetchall()

        # Building costs
        cur.execute("""
            SELECT *
            FROM building_costs
            WHERE building_id = %s AND year = %s
        """, (building_id, year))
        costs = cur.fetchall()

        # Individual costs
        cur.execute("""
            SELECT *
            FROM individual_costs
            WHERE building_id = %s AND year = %s
        """, (building_id, year))
        individual_costs = cur.fetchall()

        # Allocation rules
        cur.execute("""
            SELECT *
            FROM building_cost_allocation
            WHERE building_id = %s
        """, (building_id,))
        allocations = cur.fetchall()

    return {
        "building": building,
        "units": units,
        "tenants": tenants,
        "costs": costs,
        "individual_costs": individual_costs,
        "allocations": allocations
    }


# =========================
# BUILD MAPS
# =========================

def build_allocation_map(data):
    allocation_map = {}

    for row in data["allocations"]:
        cost_type = row["cost_type"]
        allocation_key = row["allocation_key"]

        if not cost_type or not allocation_key:
            raise ValueError(f"Invalid allocation row: {row}")

        allocation_map[cost_type] = allocation_key

    return allocation_map

def get_unit_by_id(unit_id, units):
    for u in units:
        if u["unit_id"] == unit_id:
            return u

    raise ValueError(f"Unit not found for unit_id: {unit_id}")

# =========================
# TENANT CALCULATION
# =========================

def calculate_for_tenant(tenant, data, allocation_map, year):
    result = {
        "tenant_id": tenant["tenant_id"],
        "building_id": tenant["building_id"],
        "unit_id": tenant["unit_id"],
        "lines": [],
        "total_costs": 0
    }

    occupancy_factor = calculate_occupancy_factor(tenant, year)

    # Building costs
    for cost in data["costs"]:
        line = calculate_cost_share(
            tenant,
            cost,
            data,
            allocation_map,
            occupancy_factor
        )

        if line:
            result["lines"].append(line)
            result["total_costs"] += line["amount"]

    # Individual costs
    for ic in data["individual_costs"]:
        if ic["unit_id"] == tenant["unit_id"]:
            amount = float(ic["amount"] or 0)

            result["lines"].append({
                "cost_type": ic["cost_type"],
                "allocation": "individual",
                "amount": amount
            })

            result["total_costs"] += amount

    return result


# =========================
# OCCUPANCY CALCULATION
# =========================
def get_billing_period(building, year):
    """
    Returns (start_date, end_date) based on first_billing_month
    """

    fbm = building["first_billing_month"]

    # Case 1: billing period starts in previous year
    if fbm >= 7:
        start = date(year - 1, fbm, 1)
        end = date(year, fbm, 1) - timedelta(days=1)
    else:
        start = date(year, fbm, 1)
        end = date(year + 1, fbm, 1) - timedelta(days=1)

    return start, end

def calculate_occupancy_factor(tenant, building, year):
    move_in = tenant["move_in"]
    move_out = tenant["move_out"]

    period_start, period_end = get_billing_period(building, year)

    total_months = 12.0

    # =========================
    # MOVE IN ADJUSTMENT
    # =========================
    if move_in and period_start <= move_in <= period_end:
        if move_in.day == 1:
            deduction = 0.0
        elif 2 <= move_in.day <= 15:
            deduction = 0.5
        else:
            deduction = 1.0

        # months before move_in
        months_before = (move_in.year - period_start.year) * 12 + (move_in.month - period_start.month)

        total_months -= months_before
        total_months -= deduction

    # =========================
    # MOVE OUT ADJUSTMENT
    # =========================
    if move_out and period_start <= move_out <= period_end:
        if 28 <= move_out.day <= 31:
            deduction = 0.0
        elif 14 <= move_out.day <= 27:
            deduction = 0.5
        else:
            deduction = 1.0

        # months after move_out
        months_after = (period_end.year - move_out.year) * 12 + (period_end.month - move_out.month)

        total_months -= months_after
        total_months -= deduction

    # =========================
    # Clamp (safety)
    # =========================
    if total_months < 0:
        total_months = 0

    return total_months / 12


# =========================
# COST DISTRIBUTION
# =========================

def calculate_cost_share(tenant, cost, data, allocation_map, occupancy_factor):
    cost_type = cost["cost_type"]
    total_amount = float(cost["amount"] or 0)

    allocation_key = get_allocation_key(cost_type, allocation_map)

    if allocation_key == "living_area":
        share = distribute_by_tenant_area(tenant, data)

    elif allocation_key == "total_area":
        share = distribute_by_total_area(tenant, data)

    elif allocation_key == "persons":
        share = distribute_by_persons(tenant, data)

    else:
        raise ValueError(f"Unknown allocation key: {allocation_key}")

    final_amount = total_amount * share * occupancy_factor

    return {
        "cost_type": cost_type,
        "allocation": allocation_key,
        "share": share,
        "amount": round(final_amount, 2)
    }


# =========================
# ALLOCATION METHODS
# =========================

def distribute_by_tenant_area(tenant, data):
    unit = get_unit_by_id(tenant["unit_id"], data["units"])

    total_area = data["building"]["total_tenant_area"]
    tenant_area = unit["living_area"]

    if total_area <= 0 or tenant_area <= 0:
        raise ValueError("Area smaller or equal 0")

    return tenant_area / total_area


def distribute_by_total_area(tenant, data):
    unit = get_unit_by_id(tenant["unit_id"], data["units"])

    total_area = data["building"]["total_area"]
    tenant_area = unit["living_area"]

    if total_area <= 0 or tenant_area <= 0:
        raise ValueError("Area smaller or equal 0")

    return tenant_area / total_area


def distribute_by_persons(tenant, data):
    """
    Placeholder for future implementation.
    """
    return 1.0


# =========================
# ALLOCATION LOOKUP
# =========================

def get_allocation_key(cost_type, allocation_map):
    if cost_type not in allocation_map:
        raise ValueError(f"No allocation key defined for cost_type: {cost_type}")

    return allocation_map[cost_type]


# =========================
# SPECIAL CASES
# =========================

def handle_commercial_water_adjustment(data):
    """
    Future:
    - subtract commercial water usage
    - distribute remaining costs
    """
    pass