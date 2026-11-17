"""
mis_core.py - cleaning + analytics engine for Smart MIS.
No Streamlit code in this file, so every function can be tested on its own.
"""
import io
import re
import difflib
import datetime as dt

import numpy as np
import pandas as pd

# ----------------------------------------------------------------------------
# Fields the app understands
# ----------------------------------------------------------------------------
FIELDS = {
    "order_date": {"label": "Order date", "required": True,
                   "aliases": ["order_date", "date", "purchase_date", "created_at", "order_created_at",
                               "invoice_date", "transaction_date", "sale_date", "order_dt"]},
    "order_id": {"label": "Order ID", "required": True,
                 "aliases": ["order_id", "orderid", "order_number", "order_no", "order_num",
                             "sales_order", "transaction_id", "invoice_number", "invoice_no"]},
    "product": {"label": "Product name", "required": True,
                "aliases": ["product", "product_name", "item", "item_name", "sku_name",
                            "listing_title", "product_title", "description"]},
    "quantity": {"label": "Quantity", "required": True,
                 "aliases": ["quantity", "qty", "units", "units_sold", "item_quantity",
                             "ordered_quantity", "product_quantity"]},
    "unit_price": {"label": "Unit selling price", "required": True,
                   "aliases": ["unit_price", "price", "selling_price", "sale_price", "item_price",
                               "product_price", "rate", "unit_selling_price"]},
    "order_status": {"label": "Order status", "required": False,
                     "unlocks": "Net sales (cancelled/returned excluded) and the Returns report",
                     "aliases": ["order_status", "status", "delivery_status", "fulfillment_status",
                                 "shipment_status", "order_state"]},
    "category": {"label": "Category", "required": False,
                 "unlocks": "Revenue by category",
                 "aliases": ["category", "product_category", "item_category", "department",
                             "product_type", "category_name"]},
    "region": {"label": "State / City", "required": False,
               "unlocks": "Sales by state / city",
               "aliases": ["state", "region", "city", "ship_state", "shipping_state", "delivery_state",
                           "customer_state", "ship_city", "shipping_city", "delivery_city"]},
    "channel": {"label": "Channel / Marketplace", "required": False,
                "unlocks": "Channel comparison (Amazon vs Flipkart vs website...)",
                "aliases": ["channel", "marketplace", "platform", "sales_channel", "source", "store"]},
    "customer_id": {"label": "Customer ID / name", "required": False,
                    "unlocks": "Repeat-customer report",
                    "aliases": ["customer_id", "customer", "customer_name", "buyer_id", "buyer_name",
                                "buyer", "customer_email"]},
    "unit_cost": {"label": "Cost price (per unit)", "required": False,
                  "unlocks": "Profit and margin report",
                  "aliases": ["unit_cost", "cost_price", "cost", "purchase_price", "cogs", "buying_price"]},
}
REQUIRED_FIELDS = [f for f, d in FIELDS.items() if d["required"]]
OPTIONAL_FIELDS = [f for f, d in FIELDS.items() if not d["required"]]

KNOWN_STATUSES = ["Delivered", "Shipped", "Out For Delivery", "Pending", "Cancelled", "Returned",
                  "Processing", "Confirmed", "Dispatched", "In Transit", "Refunded", "Completed",
                  "RTO", "Failed", "Undelivered"]
STATUS_ALIASES = {"canceled": "Cancelled"}
# Any status matching this is a "lost" sale: not counted in net sales.
LOST_RE = r"cancel|return|refund|\brto\b|fail|undeliver|reject|lost"


# ----------------------------------------------------------------------------
# Column detection
# ----------------------------------------------------------------------------
def clean_header(value):
    value = str(value).strip().lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    return value.strip("_")


def make_unique_headers(columns):
    result, counts = [], {}
    for col in columns:
        base = str(col).strip() or "unnamed_column"
        counts[base] = counts.get(base, 0) + 1
        result.append(base if counts[base] == 1 else f"{base}_{counts[base]}")
    return result


def detect_columns(columns):
    """Exact-alias detection. Each source column is used for at most one field."""
    normalized = {}
    for original in columns:
        normalized.setdefault(clean_header(original), original)
    detected, used = {}, set()
    for field, details in FIELDS.items():
        for alias in details["aliases"]:
            key = clean_header(alias)
            if key in normalized and normalized[key] not in used:
                detected[field] = normalized[key]
                used.add(normalized[key])
                break
    return detected


# ----------------------------------------------------------------------------
# Cleaning helpers
# ----------------------------------------------------------------------------
def parse_dates(series):
    """Real Excel dates are kept; text dates are read day-first (DD-MM-YYYY); ISO dates as ISO.
    Impossible dates (31-02-2026, TBD) become NaT."""
    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    is_real = series.map(lambda v: isinstance(v, (pd.Timestamp, dt.datetime, dt.date)))
    if is_real.any():
        out.loc[is_real] = pd.to_datetime(series[is_real], errors="coerce")
    text = series[~is_real & series.notna()].astype(str).str.strip()
    iso = text.str.match(r"^\d{4}-\d{1,2}-\d{1,2}")
    if iso.any():
        out.loc[iso[iso].index] = pd.to_datetime(text[iso], errors="coerce", format="%Y-%m-%d", exact=False)
    rest = text[~iso]
    if len(rest):
        out.loc[rest.index] = pd.to_datetime(rest, errors="coerce", dayfirst=True, format="mixed")
    return out.dt.normalize()


def normalise_status(series):
    """Tidy case/spaces, then auto-correct near-misses such as 'D3livered' -> 'Delivered'."""
    s = series.astype("string").str.strip().str.replace(r"\s+", " ", regex=True).str.title()
    lookup = {k.lower(): k for k in KNOWN_STATUSES}
    cache = {}

    def fix(v):
        if pd.isna(v) or v == "":
            return v
        if v not in cache:
            low = v.lower()
            if low in STATUS_ALIASES:
                cache[v] = STATUS_ALIASES[low]
            elif low in lookup:
                cache[v] = lookup[low]
            else:
                m = difflib.get_close_matches(low, list(lookup), n=1, cutoff=0.8)
                cache[v] = lookup[m[0]] if m else v
        return cache[v]

    return s.map(fix).astype("string")


def canonical_text(series):
    """Collapse extra spaces and merge case variants ('YOGA MAT', 'yoga mat') into the commonest spelling.
    Blank values stay missing."""
    s = series.astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
    s = s.mask((s == "").fillna(False))
    key = s.str.casefold()
    canon = s.groupby(key).agg(lambda x: x.value_counts().index[0])
    return key.map(canon).astype("string")


def to_number(series, strip_currency=True):
    txt = series.astype(str)
    if strip_currency:
        txt = txt.str.replace(r"(?i)rs\.?|inr|[₹$,\s]", "", regex=True)
    else:
        txt = txt.str.replace(r"[,\s]", "", regex=True)
    return pd.to_numeric(txt, errors="coerce")


# ----------------------------------------------------------------------------
# Main cleaning function
# ----------------------------------------------------------------------------
def process_mapped_data(source_df, mapping):
    """Returns (clean, error_rows). `mapping` = {field: source column name or None}."""
    sel = pd.DataFrame(index=source_df.index)
    mapped = [f for f in FIELDS if mapping.get(f)]
    for f in FIELDS:
        sel[f] = source_df[mapping[f]] if mapping.get(f) else pd.NA
    sel.insert(0, "source_row", range(2, len(sel) + 2))

    # ---- standard formats ----
    sel["order_date"] = parse_dates(sel["order_date"])
    sel["quantity"] = to_number(sel["quantity"], strip_currency=False)
    sel["unit_price"] = to_number(sel["unit_price"])
    sel["order_id"] = sel["order_id"].astype("string").str.strip()
    sel["product"] = canonical_text(sel["product"])

    has_status = "order_status" in mapped
    if has_status:
        sel["order_status"] = normalise_status(sel["order_status"])
        sel["is_lost"] = sel["order_status"].str.lower().str.contains(LOST_RE, regex=True).fillna(False).astype(bool)
    else:
        sel["order_status"] = pd.Series(pd.NA, index=sel.index, dtype="string")
        sel["is_lost"] = False
    for f in ("category", "region", "channel"):
        if f in mapped:
            sel[f] = canonical_text(sel[f]).fillna("Unknown")
        else:
            sel[f] = pd.Series(pd.NA, index=sel.index, dtype="string")
    if "customer_id" in mapped:
        c = sel["customer_id"].astype("string").str.strip().str.lower()
        sel["customer_id"] = c.mask((c == "").fillna(False))
    else:
        sel["customer_id"] = pd.Series(pd.NA, index=sel.index, dtype="string")
    if "unit_cost" in mapped:
        cost = to_number(sel["unit_cost"])
        sel["unit_cost"] = cost.where(cost >= 0)
    else:
        sel["unit_cost"] = np.nan

    # ---- row-level checks ----
    errors = pd.Series("", index=sel.index, dtype=object)
    warnings = pd.Series("", index=sel.index, dtype=object)

    def add(target, mask, msg):
        mask = mask.fillna(False).astype(bool)
        target.loc[mask] = target.loc[mask].map(lambda old: f"{old}; {msg}" if old else msg)

    today = pd.Timestamp.today().normalize()
    add(errors, sel["order_date"].isna(), "Missing or invalid date")
    add(errors, sel["order_date"] > today, "Order date is in the future")
    add(errors, sel["order_id"].isna() | sel["order_id"].eq(""), "Missing order ID")
    add(errors, sel["product"].isna(), "Missing product")
    if has_status:
        add(errors, sel["order_status"].isna() | sel["order_status"].eq(""), "Missing order status")
        add(warnings, ~sel["order_status"].isin(KNOWN_STATUSES) & sel["order_status"].notna()
            & sel["order_status"].ne(""), "Unrecognised order status")
    add(errors, sel["quantity"].isna(), "Missing or invalid quantity")
    add(errors, sel["quantity"].le(0), "Quantity must be greater than zero")
    add(errors, sel["quantity"].notna() & (sel["quantity"] % 1 != 0), "Quantity must be a whole number")
    add(errors, sel["unit_price"].isna(), "Missing or invalid price")
    add(errors, sel["unit_price"].lt(0), "Price cannot be negative")
    add(warnings, sel["unit_price"].eq(0), "Price is zero")

    # suspicious values: excluded so they cannot distort revenue; listed for review
    med_price = sel.groupby("product")["unit_price"].transform("median")
    add(errors, sel["unit_price"] > med_price * 10, "Suspicious price (10x above product median)")
    med_qty = sel["quantity"].median()
    add(errors, sel["quantity"] > max(50, 10 * med_qty), "Suspicious quantity (unusually large)")

    # duplicates: keep the FIRST row, send later copies to the error list
    cols = mapped
    ok = errors.eq("")
    keeper = sel[ok].groupby(cols, dropna=False)["source_row"].transform("first")
    is_exact = keeper != sel.loc[ok, "source_row"]
    errors.loc[keeper.index[is_exact]] = "Exact duplicate of source row " + keeper[is_exact].astype(str)

    ok = errors.eq("")
    keeper2 = sel[ok].groupby(["order_id", "product"])["source_row"].transform("first")
    is_conf = keeper2 != sel.loc[ok, "source_row"]
    errors.loc[keeper2.index[is_conf]] = ("Same Order ID + product as source row " + keeper2[is_conf].astype(str)
                                          + " but different details - review")

    sel["quality_issues"] = errors.astype("string")
    sel["quality_warnings"] = warnings.astype("string")
    sel["revenue"] = sel["quantity"] * sel["unit_price"]
    sel["profit"] = sel["revenue"] - sel["quantity"] * sel["unit_cost"]

    error_rows = sel[sel["quality_issues"] != ""].copy()
    clean = sel[sel["quality_issues"] == ""].copy()
    return clean, error_rows


# ----------------------------------------------------------------------------
# Analytics
# ----------------------------------------------------------------------------
WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def money(v):
    return f"₹{v:,.0f}"


def previous_window(start, end, min_date):
    """Previous comparison period. Full calendar month(s) are compared with the previous month(s);
    any other range with the previous equal-length window. None if it starts before the data does."""
    start, end = pd.Timestamp(start), pd.Timestamp(end)
    if start.day == 1 and (end + pd.Timedelta(days=1)).day == 1:
        n = (end.year - start.year) * 12 + end.month - start.month + 1
        ps = start - pd.DateOffset(months=n)
        label = "the previous month" if n == 1 else f"the previous {n} months"
    else:
        days = (end - start).days + 1
        ps = start - pd.Timedelta(days=days)
        label = f"the previous {days} days"
    pe = start - pd.Timedelta(days=1)
    if ps < pd.Timestamp(min_date):
        return None
    return ps, pe, label


def monthly_table(sales, start, end):
    m = (sales.assign(Month=sales["order_date"].dt.to_period("M")).groupby("Month")
         .agg(Revenue=("revenue", "sum"), Orders=("order_id", "nunique"), Units=("quantity", "sum")))
    idx = pd.period_range(start=pd.Timestamp(start), end=pd.Timestamp(end), freq="M")
    m = m.reindex(idx, fill_value=0)
    m["AOV"] = (m["Revenue"] / m["Orders"]).where(m["Orders"] > 0)
    m["MoM_%"] = (m["Revenue"].pct_change(fill_method=None) * 100).replace([np.inf, -np.inf], np.nan).round(1)
    m["Note"] = ["Partial month" if (p.start_time < pd.Timestamp(start) or p.end_time.normalize() > pd.Timestamp(end))
                 else "" for p in m.index]
    m = m.reset_index().rename(columns={"index": "Month"})
    m["Month"] = m["Month"].astype(str)
    m["Revenue"] = m["Revenue"].round(0)
    m["AOV"] = m["AOV"].round(0)
    return m


def weekday_table(sales, start, end):
    days = pd.Series(pd.date_range(start, end)).dt.day_name().value_counts()
    rev = sales.groupby(sales["order_date"].dt.day_name())["revenue"].sum()
    t = pd.DataFrame({"Revenue": rev, "Days_in_range": days}).reindex(WEEKDAYS).fillna(0)
    t["Avg_Revenue_per_day"] = (t["Revenue"] / t["Days_in_range"].replace(0, np.nan)).round(0)
    t["Revenue"] = t["Revenue"].round(0)
    return t.reset_index().rename(columns={"index": "Weekday"})


def product_table(rng, sales_all, end):
    """ABC class, share, cancel/return rate and 30-day trend per product."""
    sales = rng[~rng["is_lost"]]
    g = (sales.groupby("product").agg(Revenue=("revenue", "sum"), Units=("quantity", "sum"),
                                      Orders=("order_id", "nunique"))
         .sort_values("Revenue", ascending=False))
    total = g["Revenue"].sum()
    g["Share_%"] = (g["Revenue"] / total * 100) if total > 0 else 0.0
    g["Cum_%"] = g["Share_%"].cumsum()
    before = g["Cum_%"] - g["Share_%"]
    g["Class"] = np.where(before < 80, "A", np.where(before < 95, "B", "C"))
    g["Avg_Price"] = (g["Revenue"] / g["Units"]).round(0)

    lines = rng.groupby("product").size()
    lost = rng[rng["is_lost"]].groupby("product").size()
    g["Lost_%"] = (lost.reindex(g.index).fillna(0) / lines.reindex(g.index) * 100).round(1)

    end = pd.Timestamp(end)
    r_start, p_start, p_end = end - pd.Timedelta(days=29), end - pd.Timedelta(days=59), end - pd.Timedelta(days=30)
    g["Trend_%"] = np.nan
    g["Trend"] = "n/a"
    if p_start >= sales_all["order_date"].min():
        rec = sales_all[(sales_all.order_date >= r_start) & (sales_all.order_date <= end)].groupby("product")["quantity"].sum()
        prv = sales_all[(sales_all.order_date >= p_start) & (sales_all.order_date <= p_end)].groupby("product")["quantity"].sum()
        r = rec.reindex(g.index).fillna(0)
        p = prv.reindex(g.index).fillna(0)
        pct = (r - p) / p.replace(0, np.nan) * 100
        z = (r - p) / np.sqrt((r + p).replace(0, np.nan))      # simple noise test: ignore random ups and downs
        enough = p >= 20
        g["Trend_%"] = pct.where(enough).round(1)
        g["Trend"] = np.where(~enough, "n/a", np.where((pct >= 25) & (z >= 2), "Rising",
                              np.where((pct <= -25) & (z <= -2), "Falling", "Stable")))
    for c in ("Revenue", "Share_%", "Cum_%"):
        g[c] = g[c].round(1 if c != "Revenue" else 0)
    return g.reset_index()


def returns_table(rng, by="product", min_lines=10):
    """Cancel/return rate per product (or per channel / region). Flags clear outliers."""
    lines = rng.groupby(by).size().rename("Lines")
    lost_df = rng[rng["is_lost"]]
    lost = lost_df.groupby(by).size().rename("Lost_lines")
    t = pd.concat([lines, lost], axis=1).fillna(0)
    t["Lost_%"] = (t["Lost_lines"] / t["Lines"] * 100).round(1)
    t["Lost_Revenue"] = lost_df.groupby(by)["revenue"].sum().reindex(t.index).fillna(0).round(0)
    overall = (len(lost_df) / len(rng) * 100) if len(rng) else 0
    t["Flag"] = np.where((t["Lost_lines"] >= min_lines) & (t["Lost_%"] >= overall * 1.5), "High", "")
    t = t.sort_values(["Lost_%", "Lost_Revenue"], ascending=False).reset_index()
    t["Lines"] = t["Lines"].astype(int)
    t["Lost_lines"] = t["Lost_lines"].astype(int)
    return t, overall


def group_table(rng, col):
    sales = rng[~rng["is_lost"]]
    g = (sales.groupby(col).agg(Revenue=("revenue", "sum"), Orders=("order_id", "nunique"),
                                Units=("quantity", "sum")).sort_values("Revenue", ascending=False))
    total = g["Revenue"].sum()
    g["Share_%"] = (g["Revenue"] / total * 100).round(1) if total > 0 else 0.0
    lines = rng.groupby(col).size()
    lost = rng[rng["is_lost"]].groupby(col).size()
    g["Lost_%"] = (lost.reindex(g.index).fillna(0) / lines.reindex(g.index) * 100).round(1)
    g["Revenue"] = g["Revenue"].round(0)
    return g.reset_index()


def customer_summary(rng):
    s = rng[(~rng["is_lost"]) & rng["customer_id"].notna()]
    if s.empty:
        return None
    c = s.groupby("customer_id").agg(Orders=("order_id", "nunique"), Revenue=("revenue", "sum"))
    repeat = c[c["Orders"] >= 2]
    top = c.sort_values("Revenue", ascending=False).head(10).reset_index()
    top["Revenue"] = top["Revenue"].round(0)
    return {
        "customers": len(c), "repeat_customers": len(repeat),
        "repeat_rate": len(repeat) / len(c) * 100,
        "repeat_revenue_share": repeat["Revenue"].sum() / c["Revenue"].sum() * 100 if c["Revenue"].sum() else 0,
        "top": top,
    }


def profit_summary(rng):
    s = rng[(~rng["is_lost"]) & rng["unit_cost"].notna()]
    if s.empty:
        return None
    p = s.groupby("product").agg(Revenue=("revenue", "sum"), Profit=("profit", "sum"), Units=("quantity", "sum"))
    p["Margin_%"] = (p["Profit"] / p["Revenue"] * 100).round(1)
    overall = s["profit"].sum() / s["revenue"].sum() * 100 if s["revenue"].sum() else 0
    low_margin_hv = p[(p["Revenue"] >= p["Revenue"].median()) & (p["Margin_%"] < overall * 0.6)].sort_values("Margin_%")
    loss = p[p["Profit"] < 0]
    p = p.sort_values("Profit", ascending=False)
    for c in ("Revenue", "Profit"):
        p[c] = p[c].round(0)
    return {"table": p.reset_index(), "overall_margin": overall, "total_profit": s["profit"].sum(),
            "coverage": len(s) / len(rng[~rng["is_lost"]]) * 100 if len(rng[~rng["is_lost"]]) else 0,
            "low_margin_high_volume": low_margin_hv.index.tolist(), "loss_making": loss.index.tolist()}


def build_action_plan(R):
    """Rule-based, plain-language suggestions with the numbers behind them."""
    items = []

    def add(priority, area, found, action):
        items.append({"Priority": priority, "Area": area, "What we found": found, "What to do": action})

    k, q = R["kpis"], R["quality"]
    if q["error_rate"] > 5:
        add("High", "Data quality",
            f"{q['errors']:,} of {q['total']:,} rows ({q['error_rate']:.1f}%) were left out because of errors or duplicates.",
            "Fix the problem at the source (billing software / marketplace export). Use the Data Quality Report.")
    if not R["has"]["order_status"]:
        add("Medium", "Net sales",
            "Order status was not mapped, so cancelled/returned orders are counted as sales.",
            "Map an order-status column to get true net sales and the Returns report.")

    g = R["growth"]
    if g is not None:
        if g["pct"] <= -10:
            add("High", "Sales trend", f"Sales fell {abs(g['pct']):.1f}% compared with {g['label']}.",
                "Use the Products and Channel tabs to find what dropped. Check stock-outs, ad spend and competitor prices.")
        elif g["pct"] >= 10:
            add("Info", "Sales trend", f"Sales grew {g['pct']:.1f}% compared with {g['label']}.",
                "Find which products drove the growth (Products tab) and keep them in stock.")
        else:
            add("Info", "Sales trend", f"Sales are stable ({g['pct']:+.1f}%) compared with {g['label']}.",
                "To grow, work on the best-selling products and the repeat-customer offers below.")

    if R["has"]["order_status"] and k["total_orders"]:
        rate = k["lost_rate"]
        if rate > 5:
            add("High" if rate > 10 else "Medium", "Cancellations & returns",
                f"{k['lost_orders']:,} orders ({rate:.1f}%) were cancelled/returned - about {money(k['lost_revenue'])} of revenue lost.",
                "Open the Returns tab to see which products and channels cause it.")
        ret = R["returns"]
        if ret is not None:
            hot = ret[ret["Flag"] == "High"].sort_values("Lost_Revenue", ascending=False)
            for _, r in hot.head(3).iterrows():
                add("High", "Return hotspot",
                    f"{r['product']}: {r['Lost_%']:.0f}% cancelled/returned (store average {R['overall_lost']:.0f}%), {money(r['Lost_Revenue'])} lost.",
                    "Check quality, photos/description, size or fit details and packaging. Pause ads on it until fixed.")

    p = R["products"]
    if len(p) >= 5:
        a = p[p["Class"] == "A"]
        add("Medium", "Best sellers",
            f"{len(a)} of {len(p)} products bring 80% of revenue: {', '.join(a['product'].head(5))}{'...' if len(a) > 5 else ''}.",
            "Never run out of these. Protect their stock, price and listing quality first.")
        rising = p[p["Trend"] == "Rising"].sort_values("Trend_%", ascending=False).head(3)
        falling = p[p["Trend"] == "Falling"].sort_values("Trend_%").head(3)
        if len(rising):
            add("Info", "Rising products",
                "Growing in the last 30 days: " + ", ".join(f"{r['product']} ({r['Trend_%']:+.0f}%)" for _, r in rising.iterrows()) + ".",
                "Increase stock and ad spend on these while demand is rising.")
        if len(falling):
            add("Medium", "Falling products",
                "Falling in the last 30 days: " + ", ".join(f"{r['product']} ({r['Trend_%']:+.0f}%)" for _, r in falling.iterrows()) + ".",
                "Check stock-outs, price changes and reviews. Consider a discount or a bundle.")
        slow = p[p["Class"] == "C"]
        if len(slow) >= 2:
            add("Medium", "Slow movers",
                f"{len(slow)} products give only {slow['Share_%'].sum():.1f}% of revenue together, e.g. {', '.join(slow['product'].tail(5))}.",
                "Stop reordering, clear stock with offers or bundles, and use the cash for best sellers.")

    w = R["weekday"]
    if R["days"] >= 28 and w["Avg_Revenue_per_day"].notna().any():
        best, worst = w.loc[w["Avg_Revenue_per_day"].idxmax()], w.loc[w["Avg_Revenue_per_day"].idxmin()]
        add("Info", "Best days", f"{best['Weekday']} is your strongest day ({money(best['Avg_Revenue_per_day'])}/day); {worst['Weekday']} is the slowest ({money(worst['Avg_Revenue_per_day'])}/day).",
            f"Run ads and launches on {best['Weekday']}; use offers to lift {worst['Weekday']}.")

    if R["region"] is not None and len(R["region"]) >= 2:
        top = R["region"].iloc[0]
        if top["Share_%"] > 40:
            add("Medium", "Region dependence", f"{top['region']} gives {top['Share_%']:.0f}% of sales.",
                "Test ads or delivery offers in your 2nd and 3rd best regions to reduce dependence.")
    if R["channel"] is not None and len(R["channel"]) >= 2:
        top = R["channel"].iloc[0]
        if top["Share_%"] > 70:
            add("Medium", "Channel dependence", f"{top['channel']} gives {top['Share_%']:.0f}% of sales.",
                "Build a second channel (own website, WhatsApp/Instagram) so one platform cannot hurt you.")
        ch = R["channel"]
        bad = ch[(ch["Lost_%"] >= R["overall_lost"] * 1.3) & (ch["Lost_%"] > 8)]
        for _, r in bad.head(2).iterrows():
            add("Medium", "Channel returns", f"{r['channel']} has {r['Lost_%']:.0f}% cancelled/returned (average {R['overall_lost']:.0f}%).",
                "Review that platform's policies, packaging and listing accuracy.")

    cs = R["customers"]
    if cs is not None:
        if cs["repeat_rate"] < 20:
            add("Medium", "Repeat customers", f"Only {cs['repeat_rate']:.0f}% of customers ordered more than once.",
                "Send a repeat-purchase offer by WhatsApp/SMS 15-30 days after delivery.")
        else:
            add("Info", "Repeat customers", f"{cs['repeat_rate']:.0f}% of customers ordered again ({cs['repeat_revenue_share']:.0f}% of revenue).",
                "Reward these customers (early access, small discounts) to keep them.")

    pr = R["profit"]
    if pr is not None:
        if pr["loss_making"]:
            add("High", "Selling at a loss", "Loss-making products: " + ", ".join(pr["loss_making"][:5]) + ".",
                "Raise the price, cut the discount or stop selling them.")
        if pr["low_margin_high_volume"]:
            add("Medium", "Thin margins", "Big revenue but very thin margin: " + ", ".join(pr["low_margin_high_volume"][:5]) + ".",
                "Raise price slightly, renegotiate cost or reduce discounts on these.")
    order = {"High": 0, "Medium": 1, "Info": 2}
    items.sort(key=lambda d: order[d["Priority"]])
    return pd.DataFrame(items)


def build_report(clean, error_rows, mapping, start, end, total_rows):
    """Everything the dashboard and the Excel report need, for the chosen date range."""
    start, end = pd.Timestamp(start), pd.Timestamp(end)
    has = {f: bool(mapping.get(f)) for f in FIELDS}
    rng = clean[(clean["order_date"] >= start) & (clean["order_date"] <= end)].copy()
    sales_all = clean[~clean["is_lost"]]
    sales = rng[~rng["is_lost"]]

    orders = sales["order_id"].nunique()
    revenue = float(sales["revenue"].sum())
    lost = rng[rng["is_lost"]]
    total_orders = rng["order_id"].nunique()
    kpis = {"revenue": revenue, "orders": orders, "units": float(sales["quantity"].sum()),
            "aov": revenue / orders if orders else 0.0,
            "lost_orders": lost["order_id"].nunique(), "total_orders": total_orders,
            "lost_rate": lost["order_id"].nunique() / total_orders * 100 if total_orders else 0.0,
            "lost_revenue": float(lost["revenue"].sum())}

    growth = None
    pw = previous_window(start, end, clean["order_date"].min())
    if pw is not None:
        prev = float(sales_all[(sales_all["order_date"] >= pw[0]) & (sales_all["order_date"] <= pw[1])]["revenue"].sum())
        if prev > 0:
            growth = {"pct": (revenue - prev) / prev * 100, "prev_revenue": prev, "label": pw[2]}

    R = {"has": has, "rng": rng, "sales": sales, "kpis": kpis, "growth": growth,
         "days": (end - start).days + 1, "start": start, "end": end,
         "quality": {"total": total_rows, "errors": len(error_rows), "error_rate": len(error_rows) / total_rows * 100 if total_rows else 0}}
    R["monthly"] = monthly_table(sales, start, end)
    R["weekday"] = weekday_table(sales, start, end)
    R["daily"] = sales.groupby("order_date", as_index=False)["revenue"].sum()
    R["products"] = product_table(rng, sales_all, end)
    R["overall_lost"] = kpis["lost_rate"]
    if has["order_status"] and len(rng):
        ret, overall_lines = returns_table(rng, "product")
        R["returns"], R["overall_lost"] = ret, overall_lines
    else:
        R["returns"] = None
    R["category"] = group_table(rng, "category") if has["category"] else None
    R["region"] = group_table(rng, "region") if has["region"] else None
    R["channel"] = group_table(rng, "channel") if has["channel"] else None
    R["returns_channel"] = returns_table(rng, "channel", 10)[0] if has["order_status"] and has["channel"] and len(rng) else None
    R["returns_region"] = returns_table(rng, "region", 10)[0] if has["order_status"] and has["region"] and len(rng) else None
    R["customers"] = customer_summary(rng) if has["customer_id"] else None
    R["profit"] = profit_summary(rng) if has["unit_cost"] else None
    R["actions"] = build_action_plan(R)
    return R


# ----------------------------------------------------------------------------
# Excel report
# ----------------------------------------------------------------------------
def build_excel(R, error_rows):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    k = R["kpis"]
    summary = pd.DataFrame({"Metric": ["From", "To", "Net sales", "Orders (net)", "Units sold", "Average order value",
                                       "Cancelled/returned orders", "Cancel/return rate %", "Revenue lost to cancel/return",
                                       "Sales growth %", "Compared with"],
                            "Value": [R["start"].date(), R["end"].date(), round(k["revenue"], 2), k["orders"], k["units"],
                                      round(k["aov"], 2), k["lost_orders"], round(k["lost_rate"], 1), round(k["lost_revenue"], 2),
                                      round(R["growth"]["pct"], 1) if R["growth"] else "n/a",
                                      R["growth"]["label"] if R["growth"] else "n/a"]})
    sheets = {"Summary": summary, "Action Plan": R["actions"], "Monthly": R["monthly"], "Products": R["products"],
              "Weekday": R["weekday"]}
    if R["returns"] is not None:
        sheets["Returns by product"] = R["returns"]
    for name, key in (("Category", "category"), ("Region", "region"), ("Channel", "channel")):
        if R[key] is not None:
            sheets[name] = R[key]
    if R["customers"] is not None:
        sheets["Top customers"] = R["customers"]["top"]
    if R["profit"] is not None:
        sheets["Profit"] = R["profit"]["table"]
    sheets["Cleaned data"] = R["rng"].drop(columns=["quality_issues", "quality_warnings"]).assign(
        order_date=lambda d: d["order_date"].dt.date)
    if len(error_rows):
        sheets["Excluded rows"] = error_rows.assign(order_date=lambda d: d["order_date"].dt.date)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name[:31], index=False)
            ws = w.sheets[name[:31]]
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="1F4E78")
                c.alignment = Alignment(horizontal="center")
            ws.freeze_panes = "A2"
            for i, col in enumerate(df.columns, 1):
                longest = max([len(str(col))] + [len(str(v)) for v in df[col].head(200).tolist()])
                ws.column_dimensions[get_column_letter(i)].width = min(max(10, longest + 2), 70)
            if name == "Action Plan":
                for row in ws.iter_rows(min_row=2):
                    for c in row:
                        c.alignment = Alignment(wrap_text=True, vertical="top")
    return buf.getvalue()
