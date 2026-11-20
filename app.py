import io
import zlib

import pandas as pd
import plotly.express as px
import streamlit as st

import mis_core as m

MONEY_COLS = ["Revenue", "Lost_Revenue", "Avg_Price", "AOV", "Profit", "unit_price"]


def show(df):
    """Show a table without the index and with thousands separators on money columns."""
    fmt = {col: "{:,.0f}" for col in df.columns if col in MONEY_COLS}
    st.dataframe(df.style.format(fmt, na_rep="") if fmt else df, hide_index=True)


st.set_page_config(page_title="Smart MIS | Decisions for small businesses", page_icon="📊", layout="wide")

st.title("📊 Smart MIS")
st.subheader("Upload your sales file. Get clear answers on what to stock, fix and promote.")
st.caption("v2.0 • Your file is processed in memory for this session only; the app does not store it.")


# ----------------------------------------------------------------------------
# File loading (cached so the app stays fast when you change filters)
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def list_sheets(name, data):
    """Sheet names + the index of the sheet that most likely holds the orders (the one with most rows)."""
    if not name.lower().endswith((".xlsx", ".xls")):
        return [], 0
    if name.lower().endswith(".xlsx"):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(io.BytesIO(data), read_only=True)
            names = wb.sheetnames
            rows = [(wb[n].max_row or 0) for n in names]
            return names, rows.index(max(rows))
        except Exception:
            pass
    return pd.ExcelFile(io.BytesIO(data)).sheet_names, 0


@st.cache_data(show_spinner="Reading your file...")
def load_file(name, data, sheet):
    if name.lower().endswith(".csv"):
        try:
            return pd.read_csv(io.BytesIO(data), encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(io.BytesIO(data), encoding="latin-1")
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet or 0)


def quality_report_bytes(error_rows, clean, total):
    warn = clean[clean["quality_warnings"] != ""]
    cols = ["source_row", "order_date", "order_id", "product", "quantity", "unit_price", "order_status",
            "quality_issues", "quality_warnings"]
    out = pd.concat([error_rows, warn], ignore_index=True)[cols]
    out["order_date"] = out["order_date"].dt.date
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, sheet_name="Quality Issues", index=False)
        pd.DataFrame({"Metric": ["Total source rows", "Valid rows", "Rows excluded (errors/duplicates)", "Rows kept with a warning"],
                      "Count": [total, len(clean), len(error_rows), len(warn)]}).to_excel(w, sheet_name="Summary", index=False)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Upload
# ----------------------------------------------------------------------------
uploaded = st.file_uploader("Upload your sales Excel or CSV file", type=["xlsx", "xls", "csv"],
                            help="Your column names do not need to match ours - you can map them in Step 1.")
if not uploaded:
    st.info("Upload a file to begin. You can try it with demo_sales_full_messy.xlsx.")
    st.markdown("""
**Required columns:** order date, order ID, product name, quantity, unit selling price.

**Optional columns (each one unlocks more reports):** order status, category, state/city,
channel/marketplace, customer ID, cost price.

**What you get:** an Action Plan in plain language, product ranking (best sellers, slow movers),
lost-order hotspots (cancelled, returned, RTO), best days to sell, repeat customers, profit by product, and a downloadable Excel report.
""")
    st.stop()

data = uploaded.getvalue()
sheets, default_sheet = list_sheets(uploaded.name, data)
sheet = None
if len(sheets) > 1:
    sheet = st.selectbox("Which sheet contains your orders?", sheets, index=default_sheet)
elif sheets:
    sheet = sheets[0]

file_id = f"{uploaded.name}|{len(data)}|{sheet}"
if st.session_state.get("mis_file_id") != file_id:
    for key in [k for k in st.session_state if k.startswith("mis_")]:
        st.session_state.pop(key, None)
    st.session_state["mis_file_id"] = file_id

try:
    source = load_file(uploaded.name, data, sheet).copy()
    source.columns = m.make_unique_headers(source.columns)
except Exception as exc:
    st.error(f"Could not read this file: {exc}")
    st.stop()
if source.empty:
    st.error("This file has no data rows.")
    st.stop()

# ----------------------------------------------------------------------------
# Step 1 - column mapping
# ----------------------------------------------------------------------------
st.header("Step 1 — Match your columns")
detected = m.detect_columns(source.columns)
req_found = sum(1 for f in m.REQUIRED_FIELDS if f in detected)
opt_found = sum(1 for f in m.OPTIONAL_FIELDS if f in detected)
st.success(f"Detected {req_found} of {len(m.REQUIRED_FIELDS)} required columns and {opt_found} of {len(m.OPTIONAL_FIELDS)} optional columns. Please check them below.")

NONE = "— Not in my file —"
options = [NONE] + list(source.columns)
tag = zlib.crc32(file_id.encode())
with st.form("mapping_form"):
    mapping = {}
    st.markdown("**Required**")
    cols = st.columns(len(m.REQUIRED_FIELDS))
    for col, field in zip(cols, m.REQUIRED_FIELDS):
        sug = detected.get(field)
        mapping[field] = col.selectbox(m.FIELDS[field]["label"], options, index=options.index(sug) if sug in options else 0,
                                       key=f"map_{field}_{tag}")
    st.markdown("**Optional — each one you add unlocks more reports**")
    cols = st.columns(3)
    for i, field in enumerate(m.OPTIONAL_FIELDS):
        sug = detected.get(field)
        mapping[field] = cols[i % 3].selectbox(f"{m.FIELDS[field]['label']}  →  {m.FIELDS[field]['unlocks']}", options,
                                               index=options.index(sug) if sug in options else 0, key=f"map_{field}_{tag}")
    submitted = st.form_submit_button("Apply mapping and analyse", type="primary")

if submitted:
    missing = [m.FIELDS[f]["label"] for f in m.REQUIRED_FIELDS if mapping[f] == NONE]
    chosen = [v for v in mapping.values() if v != NONE]
    dup_cols = sorted({v for v in chosen if chosen.count(v) > 1})
    if missing:
        st.error("Please choose a column for: " + ", ".join(missing))
        st.stop()
    if dup_cols:
        st.error("The same column is used for more than one field: " + ", ".join(dup_cols))
        st.stop()
    mapping = {f: (None if v == NONE else v) for f, v in mapping.items()}
    with st.spinner("Cleaning and checking your data..."):
        clean, error_rows = m.process_mapped_data(source, mapping)
    st.session_state.update(mis_clean=clean, mis_errors=error_rows, mis_mapping=mapping, mis_total=len(source))
elif "mis_clean" not in st.session_state:
    st.info("Check the matches above, then click **Apply mapping and analyse**.")
    st.stop()

clean, error_rows = st.session_state["mis_clean"], st.session_state["mis_errors"]
mapping, total_rows = st.session_state["mis_mapping"], st.session_state["mis_total"]

# ----------------------------------------------------------------------------
# Step 2 - data quality
# ----------------------------------------------------------------------------
st.header("Step 2 — Data check")
warn_count = int((clean["quality_warnings"] != "").sum())
c1, c2, c3, c4 = st.columns(4)
c1.metric("Rows in file", f"{total_rows:,}")
c2.metric("Rows used", f"{len(clean):,}")
c3.metric("Rows left out", f"{len(error_rows):,}", help="Errors, blanks, impossible values and duplicate copies.")
c4.metric("Kept with a warning", f"{warn_count:,}", help="For example price is zero, or an unrecognised order status.")
if clean.empty:
    st.error("No valid rows left after cleaning. Open the list below to see why.")
with st.expander("See which rows were left out and why"):
    if len(error_rows):
        show(error_rows[["source_row", "order_id", "product", "quality_issues"]].head(1000))
        st.caption("Showing up to 1,000 rows. Download the full list below.")
    else:
        st.success("No data errors found.")
if len(error_rows) or warn_count:
    st.download_button("⬇️ Download data quality report (Excel)", quality_report_bytes(error_rows, clean, total_rows),
                       file_name="smart_mis_data_quality_report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
if clean.empty:
    st.stop()

locked = [f for f in m.OPTIONAL_FIELDS if not mapping.get(f)]
if locked:
    st.info("**Unlock more reports:** " + " • ".join(f"add **{m.FIELDS[f]['label']}** → {m.FIELDS[f]['unlocks']}" for f in locked))

# ----------------------------------------------------------------------------
# Step 3 - dashboard
# ----------------------------------------------------------------------------
st.header("Step 3 — Your business dashboard")
min_date, max_date = clean["order_date"].min().date(), clean["order_date"].max().date()
picked = st.date_input("Choose dates (click a start date, then an end date)", value=(min_date, max_date),
                       min_value=min_date, max_value=max_date, key=f"mis_range_{tag}")
if isinstance(picked, (tuple, list)):
    if len(picked) != 2:
        st.info("Please pick an END date as well to apply the filter.")
        st.stop()
    start_date, end_date = picked
else:
    start_date = end_date = picked

R = m.build_report(clean, error_rows, mapping, start_date, end_date, total_rows)
k = R["kpis"]
if R["rng"].empty:
    st.warning("There are no orders in this date range. Please choose different dates.")
    st.stop()

top = st.columns(5)
top[0].metric("Net sales", m.money(k["revenue"]))
top[1].metric("Orders", f"{k['orders']:,}")
top[2].metric("Average order value", m.money(k["aov"]))
if R["has"]["order_status"]:
    top[3].metric("Lost order rate", f"{k['lost_rate']:.1f}%", help=f"{k['lost_orders']:,} orders were lost (cancelled, returned, refunded, RTO or failed), about {m.money(k['lost_revenue'])} of revenue. Lost orders are not counted in net sales.")
else:
    top[3].metric("Units sold", f"{k['units']:,.0f}")
if R["growth"]:
    top[4].metric("Sales growth", f"{R['growth']['pct']:+.1f}%", help=f"Compared with {R['growth']['label']} ({m.money(R['growth']['prev_revenue'])})")
    st.caption(f"Growth is compared with {R['growth']['label']}.")
else:
    top[4].metric("Sales growth", "N/A", help="There is no earlier period of the same length inside your data.")
    st.caption("Sales growth needs an earlier period of the same length in your data. Try choosing one full month.")

t_act, t_over, t_prod, t_ret, t_mkt, t_profit, t_data = st.tabs(
    ["✅ What to do", "📈 Sales overview", "📦 Products", "↩️ Lost orders", "🌍 Markets & customers", "💰 Profit", "📥 Data & downloads"])

# ---- What to do ----
with t_act:
    st.subheader("Your action plan")
    st.caption("Suggestions are generated from your numbers using simple rules. Treat them as pointers to check, not guarantees.")
    acts = R["actions"]
    if acts.empty:
        st.success("Nothing urgent found for this date range.")
    for _, a in acts.iterrows():
        text = f"**{a['Area']}** — {a['What we found']}\n\n👉 {a['What to do']}"
        {"High": st.error, "Medium": st.warning}.get(a["Priority"], st.info)(text)

# ---- Overview ----
with t_over:
    st.subheader("Sales by month")
    mon = R["monthly"]
    st.plotly_chart(px.bar(mon, x="Month", y="Revenue", title="Net sales by month", labels={"Revenue": "Net sales (₹)"}))
    show(mon)
    if (mon["Note"] != "").any():
        st.caption("Months marked 'Partial month' are not fully inside the chosen dates, so compare them with care.")
    left, right = st.columns(2)
    with left:
        st.plotly_chart(px.line(R["daily"], x="order_date", y="revenue", title="Daily net sales", labels={"order_date": "Date", "revenue": "Net sales (₹)"}))
    with right:
        wd = R["weekday"]
        st.plotly_chart(px.bar(wd, x="Weekday", y="Avg_Revenue_per_day", title="Average sales per day of the week", labels={"Avg_Revenue_per_day": "Average sales per day (₹)"}))

# ---- Products ----
with t_prod:
    p = R["products"]
    st.subheader("Product ranking")
    st.caption("**Class A** = products that make the first 80% of revenue (protect these). **B** = next 15%. **C** = last 5% (slow movers). "
               "**Trend** compares units sold in the last 30 days with the 30 days before, and only shows clear changes.")
    detail_cols = ["Prev_30d_units", "Last_30d_units"]
    show(p.drop(columns=detail_cols))
    topn = p.head(10).sort_values("Revenue")
    st.plotly_chart(px.bar(topn, x="Revenue", y="product", orientation="h", title="Top 10 products by revenue", labels={"Revenue": "Net sales (₹)", "product": "Product"}))

    st.subheader("Rising and falling products")
    end_d = R["end"]
    fmt = lambda d: d.strftime("%d %b %Y")
    st.caption(f"Compares units sold in {fmt(end_d - pd.Timedelta(days=29))} – {fmt(end_d)} with {fmt(end_d - pd.Timedelta(days=59))} – {fmt(end_d - pd.Timedelta(days=30))}. "
               "A product is only called Rising or Falling when it changed by 30% or more AND the change is too big to be normal ups and downs. "
               "Every other product is 'Stable'.")
    no_data = (p["Trend"] == "n/a").all()
    a1, b1 = st.columns(2)
    for col, label, word, order in ((a1, "🔼 Rising", "Rising", False), (b1, "🔽 Falling", "Falling", True)):
        with col:
            st.markdown(f"**{label}**")
            sub = p[p["Trend"] == word][["product", "Prev_30d_units", "Last_30d_units", "Trend_%"]].sort_values("Trend_%", ascending=order)
            if not sub.empty:
                show(sub)
            elif no_data:
                st.info("Not enough data. Trends need at least 60 days of data ending on your end date.")
            else:
                st.info(f"No product is clearly {word.lower()}. The sales changes are small or within normal ups and downs.")
    with st.expander("See the 30-day change for every product (including small changes)"):
        show(p[["product", "Prev_30d_units", "Last_30d_units", "Trend_%", "Trend"]].sort_values("Trend_%", ascending=False))

# ---- Returns ----
with t_ret:
    if R["returns"] is None:
        st.info("Add your **Order status** column in Step 1 to unlock this report.")
    else:
        x1, x2, x3 = st.columns(3)
        x1.metric("Lost orders", f"{k['lost_orders']:,}")
        x2.metric("Lost order rate", f"{k['lost_rate']:.1f}%")
        x3.metric("Revenue lost", m.money(k["lost_revenue"]))
        st.caption("**Lost order** = an order that was cancelled, returned, refunded, RTO (return to origin), failed or undelivered. "
                   "Lost orders are left out of net sales. **Lost order rate** = lost orders ÷ all orders in the chosen dates.")
        if R["lost_types"] is not None:
            st.subheader("Why orders are lost")
            lt = R["lost_types"]
            lc1, lc2 = st.columns([2, 3])
            with lc1:
                show(lt)
            lc2.plotly_chart(px.bar(lt.sort_values("Orders"), x="Orders", y="Type", orientation="h", title="Lost orders by type"))
        st.subheader("By product")
        st.caption("'High' = at least 1.5× the store average with 10 or more lost order lines. 'Main_type' is the most common reason.")
        ret = R["returns"]
        show(ret)
        st.plotly_chart(px.bar(ret.head(10).sort_values("Lost_%"), x="Lost_%", y="product", orientation="h",
                               title="Top 10 products by lost-order rate (%)"))
        for name, key in (("By channel", "returns_channel"), ("By state / city", "returns_region")):
            if R[key] is not None:
                st.subheader(name)
                show(R[key])

# ---- Markets & customers ----
with t_mkt:
    for title, key, fld, unlock in (("Category", "category", "category", "Category"),
                                    ("State / city", "region", "region", "State / City"),
                                    ("Channel / marketplace", "channel", "channel", "Channel / Marketplace")):
        st.subheader(title)
        if R[key] is None:
            st.info(f"Add your **{unlock}** column in Step 1 to unlock this report.")
            continue
        g = R[key]
        left, right = st.columns([3, 2])
        with left:
            show(g)
        right.plotly_chart(px.bar(g.head(10).sort_values("Revenue"), x="Revenue", y=fld, orientation="h"))
    st.subheader("Customers")
    cs = R["customers"]
    if cs is None:
        st.info("Add your **Customer ID / name** column in Step 1 to unlock the repeat-customer report.")
    else:
        y1, y2, y3 = st.columns(3)
        y1.metric("Customers", f"{cs['customers']:,}")
        y2.metric("Repeat customers", f"{cs['repeat_rate']:.0f}%", help="Customers with 2 or more orders")
        y3.metric("Revenue from repeat customers", f"{cs['repeat_revenue_share']:.0f}%")
        st.markdown("**Top 10 customers by revenue**")
        show(cs["top"])

# ---- Profit ----
with t_profit:
    pr = R["profit"]
    if pr is None:
        st.info("Add your **Cost price (per unit)** column in Step 1 to unlock profit and margin.")
    else:
        z1, z2, z3 = st.columns(3)
        z1.metric("Estimated profit", m.money(pr["total_profit"]))
        z2.metric("Overall margin", f"{pr['overall_margin']:.1f}%")
        z3.metric("Rows with a cost price", f"{pr['coverage']:.0f}%")
        st.caption("Profit = selling price − cost price (before shipping, marketplace fees and ads). Only rows with a cost price are included.")
        pt = pr["table"]
        show(pt)
        st.plotly_chart(px.bar(pt.sort_values("Margin_%").head(15), x="Margin_%", y="product", orientation="h",
                               title="Lowest-margin products"))

# ---- Data & downloads ----
with t_data:
    st.subheader("Cleaned data for the selected dates")
    view = R["rng"].drop(columns=["quality_issues", "quality_warnings"]).copy()
    view["order_date"] = view["order_date"].dt.date
    view = view.dropna(axis=1, how="all")
    show(view.head(500))
    st.caption("Showing the first 500 rows. Downloads contain everything.")
    st.download_button("⬇️ Download cleaned data (CSV)", view.to_csv(index=False).encode("utf-8"),
                       file_name="smart_mis_cleaned_data.csv", mime="text/csv")
    st.divider()
    st.subheader("Full Excel MIS report")
    st.caption("Summary, action plan, monthly, products, lost orders, markets, customers, profit, cleaned data and excluded rows.")
    key = f"mis_xlsx_{start_date}_{end_date}"
    if st.button("Prepare Excel MIS report"):
        with st.spinner("Building the Excel report..."):
            st.session_state["mis_xlsx_key"] = key
            st.session_state["mis_xlsx"] = m.build_excel(R, error_rows)
    if st.session_state.get("mis_xlsx_key") == key and "mis_xlsx" in st.session_state:
        st.download_button("⬇️ Download Excel MIS report", st.session_state["mis_xlsx"], file_name="smart_mis_report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
