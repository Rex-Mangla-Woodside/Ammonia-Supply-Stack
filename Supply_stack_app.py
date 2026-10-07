"""
Ammonia Supply Stack - Streamlit app (run locally)
Run:   python -m streamlit run Supply_stack_app.py
Needs: pip install streamlit plotly pandas openpyxl
Keep this script in the same folder as the Excel file.

Supply : '11. Supply Stack'
         col A (no header) = country | ... | BNA FOB Cost | ... | monthly export columns
Demand : 'Total import' row on '10. S&D Model Forecasts' (taken directly,
         NOT summed from country rows). Negative values -> absolute value.
"""
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Ammonia Supply Stack", layout="wide")

# ---------------- Settings (edit if needed) ----------------
EXCEL_FILE = "Supply Stack.xlsx"
SHEET = "11. Supply Stack"                  # supply stack sheet
COST_HEADER = "bna fob cost"                # cost column header (case-insensitive)
FORECAST_SHEET = "10. S&D Model Forecasts"  # sheet with the Total import row
TOTAL_LABEL = "total import"                # label of the demand row (case-insensitive)
DEST_SHEET = None       # sheet with 'Representative destination' table; None = auto-detect
LABEL_FONT_SIZE = 16    # country label size inside bars

MONTH_FORMATS = ["%b-%y", "%b %y", "%b-%Y", "%b %Y", "%Y %b", "%Y-%b", "%B %Y", "%Y-%m"]


def to_month(h):
    """Return a Timestamp if the header is a month (Excel date or text like 'Jan-26'), else None."""
    if isinstance(h, (pd.Timestamp, datetime)):
        return pd.Timestamp(h)
    s = str(h).strip()
    for fmt in MONTH_FORMATS:
        try:
            return pd.to_datetime(s, format=fmt)
        except ValueError:
            pass
    return None


def norm(h):
    return " ".join(str(h).strip().lower().split())


@st.cache_data
def load_exports():
    """Monthly exports per country from the supply stack sheet."""
    raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET, header=None)

    # Header row = first row containing the cost header
    header_row = next((r for r in range(min(len(raw), 60))
                       if any(norm(v) == COST_HEADER for v in raw.iloc[r])), None)
    if header_row is None:
        return None, f"No '{COST_HEADER}' header found on '{SHEET}'.", []

    headers = raw.iloc[header_row].tolist()
    body = raw.iloc[header_row + 1:].reset_index(drop=True)
    i_fob = [norm(h) for h in headers].index(COST_HEADER)

    # Country = column with header 'Country' if present, else column A (blank header)
    low = [norm(h) for h in headers]
    i_country = low.index("country") if "country" in low else 0

    # Ignore old country-import columns (anything right of a header mentioning 'import')
    i_imp = next((i for i, h in enumerate(low) if "import" in h and i > i_fob), None)
    last_col = i_imp if i_imp is not None else len(headers)

    month_cols = [(i, to_month(h)) for i, h in enumerate(headers[:last_col])
                  if to_month(h) is not None]
    if not month_cols:
        return None, f"No monthly export columns found on '{SHEET}'.", [str(h) for h in headers]

    df = pd.DataFrame({
        "Country": body[i_country].astype(str).str.strip(),
        "BNA_FOB": pd.to_numeric(body[i_fob], errors="coerce"),
    })
    keep = (body[i_country].notna() & df["BNA_FOB"].notna()
            & ~df["Country"].str.lower().str.contains("total"))

    rows = [pd.DataFrame({"Country": df["Country"], "BNA_FOB": df["BNA_FOB"], "Month": m,
                          "kt": pd.to_numeric(body[i], errors="coerce").fillna(0)})[keep]
            for i, m in month_cols]
    info = (f"Supply: '{SHEET}', header row {header_row + 1}, cost col '{headers[i_fob]}', "
            f"{len(month_cols)} month cols ({headers[month_cols[0][0]]} … "
            f"{headers[month_cols[-1][0]]})")
    return pd.concat(rows, ignore_index=True), info, []


@st.cache_data
def load_total_imports():
    """Read the 'Total import' row directly from the S&D forecast sheet (monthly, kt)."""
    raw = pd.read_excel(EXCEL_FILE, sheet_name=FORECAST_SHEET, header=None)

    is_total = raw.apply(lambda col: col.astype(str).str.strip().str.lower()
                         .eq(TOTAL_LABEL)).any(axis=1)
    if not is_total.any():
        raise ValueError(f"'{TOTAL_LABEL}' row not found on '{FORECAST_SHEET}'.")
    r_total = is_total.idxmax()

    r_header = None
    for r in range(r_total - 1, -1, -1):
        if sum(to_month(v) is not None for v in raw.iloc[r]) >= 6:
            r_header = r
            break
    if r_header is None:
        raise ValueError("No month header row found above the Total import row.")

    records = []
    for c in raw.columns:
        m = to_month(raw.iat[r_header, c])
        if m is None:
            continue
        v = pd.to_numeric(raw.iat[r_total, c], errors="coerce")
        records.append({"Month": m, "kt": abs(float(v)) if pd.notna(v) else 0.0})
    return pd.DataFrame(records)


@st.cache_data
def load_destinations():
    """Return {country: [(destination, share), ...]} from the Representative destination table."""
    xl = pd.ExcelFile(EXCEL_FILE)
    sheets = [DEST_SHEET] if DEST_SHEET else xl.sheet_names
    for sh in sheets:
        raw = pd.read_excel(xl, sheet_name=sh, header=None)
        for r in range(min(len(raw), 30)):
            row = [str(v).strip().lower() for v in raw.iloc[r].tolist()]
            dest_cols = [i for i, v in enumerate(row) if v.startswith("representative destination")]
            if "row labels" not in row or not dest_cols:
                continue
            i_country = row.index("row labels")
            out = {}
            for _, rec in raw.iloc[r + 1:].iterrows():
                country = rec.iloc[i_country]
                if pd.isna(country) or str(country).strip().lower().startswith("grand total"):
                    continue
                items = []
                for i in dest_cols:
                    dest, share = rec.iloc[i], pd.to_numeric(rec.iloc[i + 1], errors="coerce")
                    if pd.notna(dest) and str(dest).strip() and pd.notna(share) and share > 0:
                        share = share * 100 if share <= 1.0 else share
                        items.append((str(dest).strip(), float(share)))
                out[str(country).strip()] = sorted(items, key=lambda x: -x[1])
            return out
    return {}


def add_period(df, basis):
    df = df.dropna(subset=["Month"]).copy()
    if basis == "Monthly":
        df["Period"] = df["Month"].dt.strftime("%b-%y")
    elif basis == "Quarterly":
        df["Period"] = ("Q" + df["Month"].dt.quarter.astype(int).astype(str) + "-"
                        + df["Month"].dt.strftime("%y"))
    else:
        df["Period"] = df["Month"].dt.year.astype(int).astype(str)
    return df


exports, supply_info, header_list = load_exports()
if exports is None:
    st.error(supply_info)
    if header_list:
        st.write("Headers Python sees on that row (send me this list):")
        st.write([f"{i}: {h}" for i, h in enumerate(header_list)])
    st.stop()
total_imports = load_total_imports()
destinations = load_destinations()

# ---------------- Sidebar parameters ----------------
st.sidebar.header("Parameters")
st.sidebar.caption(supply_info)
basis = st.sidebar.radio("Basis", ["Monthly", "Quarterly", "Annual"])
exports = add_period(exports, basis)
total_imports = add_period(total_imports, basis)

order = exports.sort_values("Month")["Period"].unique().tolist()
period = st.sidebar.selectbox("Period", order)

# ---------------- Build stack ----------------
d = (exports[exports["Period"] == period]
     .groupby(["Country", "BNA_FOB"], as_index=False)["kt"].sum()
     .rename(columns={"kt": "Export_kt"}))
d = d[d["Export_kt"] > 0].sort_values("BNA_FOB").reset_index(drop=True)
if d.empty:
    st.warning("No exports in this period.")
    st.stop()

d["x_end"] = d["Export_kt"].cumsum()
d["x_start"] = d["x_end"] - d["Export_kt"]
d["x_mid"] = d["x_start"] + d["Export_kt"] / 2
total = d["Export_kt"].sum()

# Import demand = Total import row (summed over months only for Q / annual)
demand = float(total_imports.loc[total_imports["Period"] == period, "kt"].sum())
if demand == 0:
    st.sidebar.warning(f"No Total import value found for {period}.")

default_price = float(round(d.loc[d["x_end"] >= demand, "BNA_FOB"].head(1).squeeze()
                            if (d["x_end"] >= demand).any() else d["BNA_FOB"].max()))
price = st.sidebar.number_input("USGC/BNA FOB price ($/t)", min_value=0.0,
                                value=default_price, step=5.0)

min_label = st.sidebar.slider("Only label bars wider than (kt)", 0.0,
                              float(d["Export_kt"].max()),
                              float(round(total * 0.02, 1)))
d["BarText"] = [f"<b>{c}</b><br>${v:.0f}" if w >= min_label else ""
                for c, v, w in zip(d["Country"], d["BNA_FOB"], d["Export_kt"])]


def dest_html(country):
    items = destinations.get(str(country).strip(), [])
    if not items:
        return "<br><i>No destination data</i>"
    lines = "".join(f"<br>  {dst}: {sh:.1f}%" for dst, sh in items)
    return "<br><br><b>Top exports</b> (last 4 quarters, Kpler):" + lines


missing = sorted(set(d["Country"].astype(str).str.strip()) - set(destinations))
if missing:
    with st.sidebar.expander("⚠️ No destination data for"):
        st.write(missing)

# ---------------- Chart ----------------
palette = (px.colors.qualitative.Safe + px.colors.qualitative.Pastel
           + px.colors.qualitative.Set3)
fig = go.Figure()
for i, r in d.iterrows():
    fig.add_trace(go.Bar(
        x=[r["x_mid"]], y=[r["BNA_FOB"]], width=[r["Export_kt"]],
        marker_color=palette[i % len(palette)],
        marker_line_color="white", marker_line_width=1,
        text=[r["BarText"]], textposition="inside", textangle=-90,
        insidetextanchor="middle", textfont=dict(size=LABEL_FONT_SIZE),
        hovertemplate=(f"<b>{r['Country']}</b><br>"
                       f"BNA FOB cost: ${r['BNA_FOB']:.0f}/t<br>"
                       f"Exports: {r['Export_kt']:,.0f} kt<br>"
                       f"Cumulative: {r['x_start']:,.0f}–{r['x_end']:,.0f} kt"
                       f"{dest_html(r['Country'])}"
                       "<extra></extra>"),
    ))

fig.add_vline(x=demand, line_dash="dash", line_color="red", line_width=2,
              annotation_text=f"<b>Import demand (S&D): {demand:,.0f} kt</b>",
              annotation_position="top left",
              annotation_font_color="red", annotation_font_size=14)
fig.add_hline(y=price, line_dash="dash", line_color="black", line_width=2,
              annotation_text=f"USGC/BNA FOB price: ${price:,.0f}/t",
              annotation_position="top left")

marginal = d[d["x_end"] >= demand].head(1)

fig.update_layout(
    title=f"Ammonia export supply stack – BNA FOB cost basis – {basis}: {period}",
    xaxis_title=f"Cumulative export volume (kt NH3, {basis.lower()})",
    yaxis_title="BNA FOB cost ($/t)",
    barmode="overlay", bargap=0, height=650, showlegend=False,
    uniformtext_minsize=10, uniformtext_mode="hide",
    hoverlabel=dict(font_size=13),
    xaxis=dict(range=[0, max(total, demand) * 1.08]),
    yaxis=dict(range=[0, max(d["BNA_FOB"].max(), price) * 1.08]),
)

# ---------------- Page ----------------
st.title("Ammonia Supply Stack")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total exports", f"{total:,.0f} kt")
c2.metric("Import demand (S&D)", f"{demand:,.0f} kt")
if not marginal.empty:
    m = marginal.iloc[0]
    c3.metric("Marginal supplier", m["Country"], f"${m['BNA_FOB']:.0f}/t",
              delta_color="off")
else:
    c3.metric("Marginal supplier", "Demand > supply")
c4.metric("USGC/BNA FOB price", f"${price:,.0f}/t")

st.plotly_chart(fig, use_container_width=True)

with st.expander("Data table"):
    st.dataframe(
        d[["Country", "Export_kt", "BNA_FOB", "x_start", "x_end"]]
        .rename(columns={"Export_kt": "Exports (kt)", "BNA_FOB": "BNA FOB ($/t)",
                         "x_start": "Cum. start (kt)", "x_end": "Cum. end (kt)"})
        .round(1),
        use_container_width=True)

with st.expander("Import demand – Total import row (S&D Model Forecasts)"):
    st.dataframe(total_imports.assign(Month=total_imports["Month"].dt.strftime("%b-%y"))
                 .round(1), use_container_width=True)