from datetime import date, timedelta, datetime
from psycopg2.extras import RealDictCursor
from yearly_import import get_connection


# =========================
# MAIN ENTRY
# =========================

def calculate_billing(building_id: str, year: int):
    conn = get_connection()

    try:
        # load data
        data = load_data(conn, building_id, year)

        # build lookup maps and do precalculations
        occupancy_map = build_occupancy_map(data, year)
        allocation_map = build_allocation_map(data)
        people_map = build_people_map(data, occupancy_map)
        special_cases = calculate_special_costs(data)

        maps = {
            "occupancy": occupancy_map,
            "allocation": allocation_map,
            "people": people_map,
            "special": special_cases
        }

        results = []

        for tenant in data["tenants"]:
            if tenant["tenant_id"] in occupancy_map:
                result = calculate_for_tenant(
                    tenant,
                    data,
                    maps
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
        "allocations": allocations,
        "year": year
    }


# =========================
# BUILD MAPS
# =========================

def build_occupancy_map(data, year):
    occupancy_map = {}

    for tenant in data["tenants"]:
        months = calculate_occupancy_months(tenant, data["building"], year)
        if months > 0:
            occupancy_map[tenant["tenant_id"]] = months

    return occupancy_map


def build_allocation_map(data):
    allocation_map = {}

    for row in data["allocations"]:
        cost_type = row["cost_type"]
        allocation_key = row["allocation_key"]

        if not cost_type or not allocation_key:
            raise ValueError(f"Invalid allocation row: {row}")

        allocation_map[cost_type] = allocation_key

    return allocation_map


def build_people_map(data, occupancy_map):
    person_map = {}
    needed = False
    for allocation in data["allocations"]:
        if allocation["allocation_key"] == "Personen":
            needed = True

    if needed:
        for tenant in data["tenants"]:
            if tenant["tenant_id"] in occupancy_map:
                person_map[tenant["tenant_id"]] = tenant["pers_count"]

    return person_map


# precalculation for distribution type "Fläche * +"
def calculate_special_costs(data):
    special_costs = {}

    for allocation in data["allocations"]:
        if allocation["allocation_key"] == "Fläche * +":
            cost_type = allocation["cost_type"]
            total_cost = 0

            for cost in data["costs"]:
                if cost["cost_type"] == cost_type:
                    total_cost = cost["amount"]

            if total_cost == 0:
                raise ValueError(cost_type + "not found. Error")

            for ic in data["individual_costs"]:
                if ic["cost_type"] == cost_type:
                    total_cost -= ic["amount"]

            special_costs[cost_type] = total_cost

    return special_costs


# =========================
# TENANT CALCULATION
# =========================

def calculate_for_tenant(tenant, data, maps):
    months = maps["occupancy"][tenant["tenant_id"]]

    result = {
        "tenant_id": tenant["tenant_id"],
        "building_id": tenant["building_id"],
        "unit_id": tenant["unit_id"],
        "months": months,
        "lines": [],
        "total_tenant_costs": 0
    }

    # Building costs
    for cost in data["costs"]:
        line = calculate_cost_share(
            tenant,
            cost,
            data,
            maps,
            months
        )

        if line:
            result["lines"].append(line)
            result["total_tenant_costs"] += line["amount"]

    # Individual costs
    for ic in data["individual_costs"]:
        if ic["unit_id"] == tenant["unit_id"]:

            # check if it is the right tenant for cases of change of tenants
            cost_date = ic.get("date")
            if not cost_date or is_date_in_tenancy_period(tenant, cost_date):
                amount = float(ic["amount"] or 0)

                result["lines"].append({
                    "type": "individual",
                    "cost_type": ic["cost_type"],
                    "allocation": ic["allocation_key"],
                    "amount": amount,
                    "total_amount": ic["total_amount"],
                    "usage": ic["usage"],
                    "price": ic["price"]
                })

                result["total_tenant_costs"] += amount

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


def calculate_occupancy_months(tenant, building, year):
    move_in = tenant["move_in"]
    move_out = tenant["move_out"]

    period_start, period_end = get_billing_period(building, year)

    if move_in > period_end or (move_out and move_out < period_start):
        return 0

    total_months = 12.0

    # =========================
    # MOVE IN ADJUSTMENT
    # =========================
    if move_in and period_start <= move_in <= period_end:

        deduction = (move_in.day - 1) / 30

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
        else:
            deduction = (30 - move_out.day) / 30

        # months after move_out
        months_after = (period_end.year - move_out.year) * 12 + (period_end.month - move_out.month)

        total_months -= months_after
        total_months -= deduction

    # =========================
    # Clamp (safety)
    # =========================
    if total_months < 0:
        total_months = 0

    return total_months


# =========================
# COST DISTRIBUTION
# =========================

def calculate_cost_share(tenant, cost, data, maps, months):
    if data["building"]["is_single_unit"]:
        return {
            "type": "general",
            "cost_type": cost["cost_type"],
            "allocation": "",
            "total_amount": "",
            "amount": round(cost["amount"], 2),
            "special_amount": ""
        }

    else:
        cost_type = cost["cost_type"]
        total_amount = float(cost["amount"] or 0)

        # initialize variables
        special_amount = 0

        allocation_key = get_allocation_key(cost_type, maps["allocation"])

        if allocation_key == "Fläche *":
            amount = distribute_by_tenant_area(tenant, data) * total_amount * (months / 12)

        elif allocation_key == "Fläche":
            amount = distribute_by_total_area(tenant, data) * total_amount * (months / 12)

        elif allocation_key == "Personen":
            amount = distribute_by_people(tenant, data, maps) * total_amount

        elif allocation_key == "Wohnungen":
            amount = distribute_by_units(tenant, data) * total_amount * (months / 12)

        elif allocation_key == "Garagen":
            amount = distribute_by_garages(tenant, data) * total_amount * (months / 12)

        elif allocation_key == "Fläche * +":
            special_amount = maps["special"][cost_type]
            amount = distribute_by_tenant_area(tenant, data) * special_amount * (months / 12)

        else:
            raise ValueError(f"Unknown allocation key: {allocation_key}")


        return {
            "type": "general",
            "cost_type": cost_type,
            "allocation": allocation_key,
            "total_amount": total_amount,
            "amount": round(amount, 2),
            "special_amount": special_amount
        }


# =========================
# ALLOCATION METHODS
# =========================

def distribute_by_tenant_area(tenant, data):
    unit = get_unit_by_id(tenant["unit_id"], data["units"])
    if unit.get("is_shop"):
        return 0
    else:
        unit = get_unit_by_id(tenant["unit_id"], data["units"])

        total_area = data["building"]["total_tenant_area"]
        tenant_area = unit["area"]

        if total_area <= 0 or tenant_area <= 0:
            raise ValueError("Area smaller or equal 0")

        return tenant_area / total_area


def distribute_by_total_area(tenant, data):
    unit = get_unit_by_id(tenant["unit_id"], data["units"])

    total_area = data["building"]["total_area"]
    tenant_area = unit["area"]

    if total_area <= 0 or tenant_area <= 0:
        raise ValueError("Area smaller or equal 0")

    return tenant_area / total_area


def distribute_by_people(tenant, data, maps):
    if "people" not in maps["special"]:
        distribute_by_people_help(maps)

    tenant_id = tenant["tenant_id"]
    tenant_people_months = maps["occupancy"][tenant_id] * maps["people"][tenant_id]
    total_people_months = maps["special"]["people"]

    return tenant_people_months / total_people_months


def distribute_by_people_help(maps):
    people_months = 0
    for tenant_id in maps["occupancy"]:
        people_months += maps["occupancy"][tenant_id] * maps["people"][tenant_id]
    maps["special"]["people"] = people_months


def distribute_by_units(tenant, data):
    total_units = data["building"]["unit_count"]

    return 1 / total_units


def distribute_by_garages(tenant, data):
    total_garages = data["building"]["gar_count"]
    tenant_garages = tenant["gar_count"]

    return tenant_garages / total_garages


# =========================
# LOOKUPS
# =========================

def get_allocation_key(cost_type, allocation_map):
    if cost_type not in allocation_map:
        raise ValueError(f"No allocation key defined for cost_type: {cost_type}")

    return allocation_map[cost_type]


def get_unit_by_id(unit_id, units):
    for u in units:
        if u["unit_id"] == unit_id:
            return u

    raise ValueError(f"Unit not found for unit_id: {unit_id}")

def is_date_in_tenancy_period(tenant, date):
    move_in = tenant.get("move_in")
    move_out = tenant.get("move_out") or datetime.max

    return move_in <= date <= move_out
