"""
MiiR Routing Assistant — Streamlit app (v2)
Run: streamlit run app.py
"""

import html
import math
import re
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    import pgeocode
    PGEOCODE_AVAILABLE = True
except ImportError:
    PGEOCODE_AVAILABLE = False

TEMPLATE_FILE = Path(__file__).parent / "MiiR_Printer_Availability_Template_v6.xlsx"
CACHE_TTL = 60  # seconds

CAP_ORDER = {"YES": 0, "YES*": 1, "LIMITED": 2, "UNTESTED": 3, "NO": 4}
ELIGIBLE_CAPS = ("YES", "YES*")
ANY_CAPABILITY_CAPS = ("YES", "YES*", "LIMITED")

NO_DECORATOR_MSG = (
    "No decorator in MiiR's network can handle this product + decoration combination. "
    "Contact operations to discuss alternative solutions."
)

CELL_STYLES = {
    "YES": "background-color:#d4edda;color:#155724;font-weight:600",
    "YES*": "background-color:#fff3cd;color:#856404;font-weight:600",
    "LIMITED": "background-color:#ffe5b4;color:#7a4a00",
    "UNTESTED": "background-color:#e2e3e5;color:#383d41",
    "NO": "background-color:#f8d7da;color:#721c24",
}


# --------------------------------------------------------------------------- #
# Data loading
# --------------------------------------------------------------------------- #

@st.cache_data(ttl=CACHE_TTL)
def load_template(mtime: float) -> dict:
    """Load all relevant sheets into dicts/dataframes."""
    xl = pd.ExcelFile(TEMPLATE_FILE, engine="openpyxl")

    def read(sheet, header_row=3):
        return xl.parse(sheet, header=header_row)

    # ---- HOLIDAYS ----
    holidays_raw = read("HOLIDAYS")
    holidays = set()
    for _, r in holidays_raw.iterrows():
        raw = r.get("Date (YYYY-MM-DD)")
        if raw is None or (isinstance(raw, float) and math.isnan(raw)):
            continue
        try:
            holidays.add(pd.to_datetime(str(raw)).date())
        except (ValueError, TypeError):
            continue

    # ---- PRODUCTS ----
    products_raw = read("PRODUCTS")
    products_by_family = {}
    for _, r in products_raw.iterrows():
        product = str(r.get("Product", "") or "").strip()
        family = str(r.get("Family", "") or "").strip()
        if not product or not family or product == "Product":
            continue
        products_by_family.setdefault(family, []).append(product)
    families = sorted(products_by_family.keys())

    # ---- DECORATORS ----
    dec_raw = read("DECORATORS")
    decorators = {}
    for _, r in dec_raw.iterrows():
        name = str(r.get("Decorator", "")).strip()
        if name and name != "Decorator" and not name.startswith("["):
            decorators[name] = {
                "cost_tier": _clean_str(r.get("Cost Tier (Low/Mid/High)"), "—"),
                "region": _clean_str(r.get("Region"), "—"),
                "int_ext": _clean_str(r.get("Internal / External"), "—"),
                "notes": _clean_str(r.get("Notes"), ""),
                "zip": _parse_zip(r.get("Zip Code")),
            }

    # ---- DECORATION_TYPES ----
    dt_raw = read("DECORATION_TYPES")
    deco_types = {}
    for _, r in dt_raw.iterrows():
        name = str(r.get("Name", "")).strip()
        if name and not name.startswith("[") and name != "Name":
            lead_raw = _clean_str(r.get("Lead Time (days)"), "")
            lead_days = _parse_lead_time(lead_raw)
            deco_types[name] = {"lead_raw": lead_raw, "lead_days": lead_days}

    # ---- CAPABILITY ----
    cap_raw = read("CAPABILITY")
    decorator_cols = [
        c for c in cap_raw.columns
        if c not in ("Product", "Decoration Type", "Notes / Caveats", "Last Updated", "Updated By")
        and not str(c).startswith("Unnamed")
        and c != "MPIX"  # MPIX never surfaces as a routing option
    ]
    capability = {}
    for _, r in cap_raw.iterrows():
        product = _clean_str(r.get("Product"), "")
        deco = _clean_str(r.get("Decoration Type"), "")
        if not product or not deco or product == "Product":
            continue
        key = (product, deco)
        caps = {}
        for col in decorator_cols:
            raw = r.get(col)
            if raw is None or (isinstance(raw, float) and math.isnan(raw)):
                caps[col] = None
                continue
            val = str(raw).strip().upper()
            caps[col] = val if val in ("YES", "YES*", "NO", "LIMITED", "UNTESTED") else None
        caveat = _clean_str(r.get("Notes / Caveats"), "")
        capability[key] = {"caps": caps, "caveat": caveat}

    return {
        "families": families,
        "products_by_family": products_by_family,
        "decorators": decorators,
        "deco_types": deco_types,
        "capability": capability,
        "holidays": holidays,
        "mtime": mtime,
    }


def _clean_str(raw, default: str) -> str:
    """Stringify a cell value, treating None/NaN/blank as `default`."""
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return default
    s = str(raw).strip()
    return s or default


def _parse_zip(raw) -> str | None:
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return None
    try:
        return f"{int(float(raw)):05d}"
    except (ValueError, TypeError):
        digits = re.sub(r"\D", "", str(raw))
        return digits or None


def _parse_lead_time(raw: str) -> int | None:
    """Extract the upper-bound lead time in days. None if no digits found ('TBD')."""
    nums = re.findall(r"\d+", raw)
    if not nums:
        return None
    return max(int(n) for n in nums)


def get_data() -> dict:
    mtime = TEMPLATE_FILE.stat().st_mtime if TEMPLATE_FILE.exists() else 0.0
    return load_template(mtime)


# --------------------------------------------------------------------------- #
# Business-day lead time
# --------------------------------------------------------------------------- #

def add_business_days(start_date: date, n_days: int, holidays: set) -> date:
    """Step forward n_days business days from start_date, skipping weekends and holidays."""
    d = start_date
    added = 0
    while added < n_days:
        d += timedelta(days=1)
        if d.weekday() < 5 and d not in holidays:
            added += 1
    return d


def compute_ship_date(lead_days, holidays: set):
    if lead_days is None:
        return None
    return add_business_days(date.today(), lead_days, holidays)


def format_ship_date(ship_date) -> str:
    if ship_date is None:
        return "TBD"
    return f"{ship_date:%B} {ship_date.day}, {ship_date:%Y}"


# --------------------------------------------------------------------------- #
# Zip-code proximity
# --------------------------------------------------------------------------- #

@st.cache_resource
def get_geocoder():
    if not PGEOCODE_AVAILABLE:
        return None
    try:
        return pgeocode.Nominatim("us")
    except Exception:
        return None


def _zip_latlon(nomi, zip_code):
    if not nomi or not zip_code:
        return None
    try:
        rec = nomi.query_postal_code(zip_code)
        lat, lon = rec.latitude, rec.longitude
        if lat is None or lon is None or math.isnan(lat) or math.isnan(lon):
            return None
        return (lat, lon)
    except Exception:
        return None


def haversine_miles(lat1, lon1, lat2, lon2) -> float:
    r_km = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlambda / 2) ** 2
    km = 2 * r_km * math.asin(math.sqrt(a))
    return km * 0.621371


def distance_miles(nomi, zip_a, zip_b):
    """Rounded integer miles between two zips, or None if either can't be geocoded."""
    loc_a = _zip_latlon(nomi, zip_a)
    loc_b = _zip_latlon(nomi, zip_b)
    if not loc_a or not loc_b:
        return None
    return round(haversine_miles(loc_a[0], loc_a[1], loc_b[0], loc_b[1]))


def is_valid_zip(zip_str: str) -> bool:
    return bool(re.fullmatch(r"\d{5}", zip_str.strip()))


# --------------------------------------------------------------------------- #
# Routing logic
# --------------------------------------------------------------------------- #

def deco_types_for_product(data: dict, product: str) -> list:
    """Decoration types that have at least one CAPABILITY row for this product."""
    return sorted({deco for (p, deco) in data["capability"] if p == product})


def matrix_columns(data: dict, product: str, deco_types: list) -> list:
    """Decorators (excl. MPIX) with at least one non-blank cell for this product."""
    cols = set()
    for deco in deco_types:
        caps = data["capability"].get((product, deco), {}).get("caps", {})
        for dec_name, val in caps.items():
            if val:
                cols.add(dec_name)
    return sorted(cols)


def product_has_any_capability(data: dict, product: str, deco_types: list) -> bool:
    """True if any (non-MPIX) decorator is YES / YES* / LIMITED anywhere for this product."""
    for deco in deco_types:
        caps = data["capability"].get((product, deco), {}).get("caps", {})
        if any(v in ANY_CAPABILITY_CAPS for v in caps.values()):
            return True
    return False


def eligible_decorators_for(data: dict, product: str, deco: str) -> list:
    """YES / YES* decorators (excl. MPIX) for a given product+decoration combo."""
    entry = data["capability"].get((product, deco), {})
    caps = entry.get("caps", {})
    caveat = entry.get("caveat", "")
    rows = []
    for dec_name, val in caps.items():
        if val not in ELIGIBLE_CAPS:
            continue
        info = data["decorators"].get(dec_name, {})
        rows.append({
            "decorator": dec_name,
            "capability": val,
            "caveat": caveat,
            "cost_tier": info.get("cost_tier", "—"),
            "region": info.get("region", "—"),
            "zip": info.get("zip"),
        })
    return rows


def sort_eligible(rows: list, nomi, customer_zip: str) -> list:
    if customer_zip:
        for r in rows:
            r["distance"] = distance_miles(nomi, r["zip"], customer_zip)
        rows.sort(key=lambda r: (r["distance"] is None, r["distance"] if r["distance"] is not None else 0))
    else:
        for r in rows:
            r["distance"] = None
        rows.sort(key=lambda r: (CAP_ORDER.get(r["capability"], 99), r["decorator"]))
    return rows


# --------------------------------------------------------------------------- #
# Matrix rendering (HTML, with hover tooltips)
# --------------------------------------------------------------------------- #

def render_matrix_html(data: dict, product: str, deco_types: list, columns: list) -> str:
    header_style = "padding:6px 10px;border:1px solid #ddd;background:#f8f9fa;color:#212529;"
    header_cells = "".join(
        f"<th style='{header_style}'>{html.escape(c)}</th>"
        for c in columns
    )
    thead = (
        f"<tr><th style='text-align:left;{header_style}'>"
        f"Decoration Type</th>{header_cells}</tr>"
    )

    body_rows = []
    for deco in deco_types:
        entry = data["capability"].get((product, deco), {})
        caps = entry.get("caps", {})
        caveat = entry.get("caveat", "")
        cells = (
            f"<td style='padding:6px 10px;border:1px solid #ddd;font-weight:600;'>{html.escape(deco)}</td>"
        )
        for col in columns:
            val = caps.get(col)
            style = CELL_STYLES.get(val, "")
            display = val if val else "—"
            title_attr = f' title="{html.escape(caveat)}"' if val and caveat else ""
            cells += (
                f"<td style='padding:6px 10px;border:1px solid #ddd;text-align:center;{style}'{title_attr}>"
                f"{html.escape(display)}</td>"
            )
        body_rows.append(f"<tr>{cells}</tr>")

    return (
        "<table style='border-collapse:collapse;width:100%;font-size:0.9rem;'>"
        f"<thead>{thead}</thead><tbody>{''.join(body_rows)}</tbody></table>"
    )


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #

st.set_page_config(page_title="MiiR Routing Assistant", page_icon="🧭", layout="centered")

st.title("🧭 MiiR Routing Assistant")
st.caption("Select a silhouette to see its capability matrix, lead times, and eligible decorators.")

if not TEMPLATE_FILE.exists():
    st.error(f"Template not found: `{TEMPLATE_FILE.name}`. Place it in the same folder as this app.")
    st.stop()

data = get_data()
nomi = get_geocoder()

# ---- Silhouette / Size inputs ----
col1, col2 = st.columns(2)
with col1:
    family = st.selectbox("Silhouette", ["— select —"] + data["families"])

product = None
if family != "— select —":
    variants = data["products_by_family"][family]
    if len(variants) > 1:
        with col2:
            product = st.selectbox("Size / Variant", ["— select —"] + variants)
        if product == "— select —":
            product = None
    else:
        product = variants[0]

customer_zip = st.text_input(
    "Customer shipping zip code (optional)",
    placeholder="e.g. 98004",
)
st.caption(
    "Enter the customer's shipping zip to see decorators sorted by proximity. "
    "Freight cost is a function of decorator-to-customer distance, so the closest "
    "decorator is usually the cheapest to ship from."
)

zip_valid = False
if customer_zip:
    if is_valid_zip(customer_zip):
        zip_valid = True
    else:
        st.caption("⚠️ Enter a valid 5-digit US zip code to sort by proximity.")
if customer_zip and zip_valid and nomi is None:
    st.caption("⚠️ Zip lookup is unavailable right now — showing capability-tier order instead.")

if not product:
    st.info("Select a silhouette above to see its capability matrix.")
    st.stop()

customer_zip_clean = customer_zip.strip() if zip_valid else None

# ---- Capability matrix ----
st.subheader(f"Capability matrix for: {product}")

deco_types = deco_types_for_product(data, product)
columns = matrix_columns(data, product, deco_types)

if not deco_types or not product_has_any_capability(data, product, deco_types):
    st.warning(NO_DECORATOR_MSG)
else:
    st.html(render_matrix_html(data, product, deco_types, columns))
    st.caption("Hover any cell for caveats. — means no capability entry on file.")

# ---- Decoration Type Details ----
if deco_types:
    st.divider()
    st.subheader("Decoration Type Details")
    selected_deco = st.selectbox("Decoration Type", deco_types, key="deco_detail")

    dt_info = data["deco_types"].get(selected_deco, {})
    lead_raw = dt_info.get("lead_raw", "—")
    lead_days = dt_info.get("lead_days")
    ship_date = compute_ship_date(lead_days, data["holidays"])

    c1, c2 = st.columns(2)
    c1.metric("Lead Time", f"{lead_raw} days" if lead_raw else "—")
    c2.metric("Suggested Ship Date", format_ship_date(ship_date))
    st.caption("Ship date skips weekends and MiiR-observed holidays, using the upper bound of any lead-time range.")

    eligible = eligible_decorators_for(data, product, selected_deco)
    if not eligible:
        st.warning(NO_DECORATOR_MSG)
    else:
        eligible = sort_eligible(eligible, nomi, customer_zip_clean)
        proximity_note = f" (sorted by proximity to {customer_zip_clean})" if customer_zip_clean else ""
        st.markdown(f"**Eligible decorators{proximity_note}:**")
        for i, r in enumerate(eligible, start=1):
            dist_str = ""
            if customer_zip_clean:
                dist_str = f" — {r['distance']} mi" if r["distance"] is not None else " — distance unknown"
            zip_str = f" ({r['zip']})" if r["zip"] else ""
            st.markdown(f"**{i}. {r['decorator']}{zip_str}**{dist_str} — {r['capability']}")
            st.caption(
                f"Cost tier: {r['cost_tier']} · Region: {r['region']}"
                + (f"\n\nCaveat: {r['caveat']}" if r["caveat"] else "")
            )

# ---- Footer ----
mtime = TEMPLATE_FILE.stat().st_mtime
last_updated = pd.Timestamp(mtime, unit="s").strftime("%Y-%m-%d %H:%M")
st.divider()
st.caption(f"Last template update: {last_updated} · Template: `{TEMPLATE_FILE.name}`")
