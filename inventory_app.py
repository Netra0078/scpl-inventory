import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# ─────────────────────────────────────────
# DATABASE SETUP
# ─────────────────────────────────────────

DB_PATH = "inventory.db"

def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS items (
            item_code TEXT PRIMARY KEY,
            description TEXT,
            unit TEXT,
            top_group TEXT,
            item_group TEXT,
            hsn TEXT,
            reorder_level REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS grn (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_no TEXT,
            document_date TEXT,
            supplier TEXT,
            item_code TEXT,
            receive_qty REAL DEFAULT 0,
            rejected_qty REAL DEFAULT 0,
            rate REAL DEFAULT 0,
            amount REAL DEFAULT 0,
            business_unit TEXT,
            warehouse TEXT,
            work_order TEXT,
            document_type TEXT
        );

        CREATE TABLE IF NOT EXISTS issues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_no TEXT,
            document_date TEXT,
            item_code TEXT,
            quantity REAL DEFAULT 0,
            rate REAL DEFAULT 0,
            amount REAL DEFAULT 0,
            business_unit TEXT,
            warehouse TEXT,
            activity TEXT,
            contractor TEXT,
            employee_name TEXT,
            document_type TEXT
        );

        CREATE TABLE IF NOT EXISTS reorder_levels (
            item_code TEXT PRIMARY KEY,
            reorder_level REAL DEFAULT 0
        );
    """)
    conn.commit()
    conn.close()

# ─────────────────────────────────────────
# PHASE 1 — EXCEL PARSING
# ─────────────────────────────────────────

def parse_grn_excel(file) -> pd.DataFrame:
    raw = pd.read_excel(file, skiprows=9, header=0)
    raw.columns = raw.iloc[0]
    df = raw.iloc[1:].reset_index(drop=True)
    df.columns.name = None
    df = df.dropna(subset=["Document No", "Item Code"])
    for col in ["Quantity", "Rate", "Amount", "Receive Quantity", "Rejected Quantity"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Document Date"] = pd.to_datetime(df["Document Date"], errors="coerce")
    df["Document No"]   = df["Document No"].astype(str).str.strip()
    df["Item Code"]     = df["Item Code"].astype(str).str.strip()
    return df

def parse_issues_excel(file) -> pd.DataFrame:
    raw = pd.read_excel(file, skiprows=9, header=0)
    raw.columns = raw.iloc[0]
    df = raw.iloc[1:].reset_index(drop=True)
    df.columns.name = None
    df = df.dropna(subset=["Document No", "Item Code"])
    for col in ["Quantity", "Rate", "Amount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Document Date"] = pd.to_datetime(df["Document Date"], errors="coerce")
    df["Document No"]   = df["Document No"].astype(str).str.strip()
    df["Item Code"]     = df["Item Code"].astype(str).str.strip()
    return df

# ─────────────────────────────────────────
# PHASE 2 — LOAD INTO SQLITE
# ─────────────────────────────────────────

def load_grn_to_db(df: pd.DataFrame):
    conn = get_conn()
    loaded = 0
    for _, row in df.iterrows():
        item_code = str(row.get("Item Code", "")).strip()
        if not item_code:
            continue
        conn.execute("""
            INSERT OR IGNORE INTO items (item_code, description, unit, top_group, item_group, hsn)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (item_code,
              str(row.get("Description", "")),
              str(row.get("Unit", "")),
              str(row.get("Top Group", "")),
              str(row.get("Item Group", "")),
              str(row.get("HSN", ""))))
        conn.execute("""
            INSERT INTO grn (document_no, document_date, supplier, item_code,
                             receive_qty, rejected_qty, rate, amount,
                             business_unit, warehouse, work_order, document_type)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (str(row.get("Document No", "")),
              str(row.get("Document Date", ""))[:10],
              str(row.get("Supplier", "")),
              item_code,
              float(row.get("Receive Quantity", 0) or 0),
              float(row.get("Rejected Quantity", 0) or 0),
              float(row.get("Rate", 0) or 0),
              float(row.get("Amount", 0) or 0),
              str(row.get("Business Unit", "")),
              str(row.get("Warehouse", "")),
              str(row.get("Work Order", "")),
              str(row.get("Document Type", ""))))
        loaded += 1
    conn.commit()
    conn.close()
    return loaded

def load_issues_to_db(df: pd.DataFrame):
    conn = get_conn()
    loaded = 0
    for _, row in df.iterrows():
        item_code = str(row.get("Item Code", "")).strip()
        if not item_code:
            continue
        conn.execute("""
            INSERT INTO issues (document_no, document_date, item_code, quantity,
                                rate, amount, business_unit, warehouse,
                                activity, contractor, employee_name, document_type)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (str(row.get("Document No", "")),
              str(row.get("Document Date", ""))[:10],
              item_code,
              float(row.get("Quantity", 0) or 0),
              float(row.get("Rate", 0) or 0),
              float(row.get("Amount", 0) or 0),
              str(row.get("Business Unit", "")),
              str(row.get("Warehouse", "")),
              str(row.get("Activity", "")),
              str(row.get("Contractor", "")),
              str(row.get("Employee Name", "")),
              str(row.get("Document Type", ""))))
        loaded += 1
    conn.commit()
    conn.close()
    return loaded

# ─────────────────────────────────────────
# PHASE 3 — STOCK BALANCE ENGINE
# ─────────────────────────────────────────

def get_stock_summary() -> pd.DataFrame:
    conn = get_conn()
    query = """
        SELECT
            i.item_code,
            i.description,
            i.unit,
            i.top_group,
            i.item_group,
            COALESCE(g.total_received, 0) AS total_received,
            COALESCE(s.total_issued, 0)   AS total_issued,
            COALESCE(g.total_received, 0) - COALESCE(s.total_issued, 0) AS current_stock,
            COALESCE(g.total_grn_amount, 0) AS total_grn_amount,
            COALESCE(rl.reorder_level, 0) AS reorder_level
        FROM items i
        LEFT JOIN (
            SELECT item_code,
                   SUM(receive_qty - rejected_qty) AS total_received,
                   SUM(amount) AS total_grn_amount
            FROM grn
            WHERE document_type NOT LIKE '%ILTO%'
            GROUP BY item_code
        ) g ON i.item_code = g.item_code
        LEFT JOIN (
            SELECT item_code, SUM(quantity) AS total_issued
            FROM issues
            GROUP BY item_code
        ) s ON i.item_code = s.item_code
        LEFT JOIN reorder_levels rl ON i.item_code = rl.item_code
        ORDER BY current_stock DESC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_bu_summary() -> pd.DataFrame:
    conn = get_conn()
    query = """
        SELECT business_unit,
               COUNT(DISTINCT item_code) AS unique_items,
               SUM(quantity) AS total_issued_qty,
               SUM(amount) AS total_issued_value
        FROM issues
        WHERE business_unit != 'nan'
        GROUP BY business_unit
        ORDER BY total_issued_value DESC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_monthly_trend() -> pd.DataFrame:
    conn = get_conn()
    query = """
        SELECT substr(document_date, 1, 7) AS month,
               SUM(quantity) AS issued_qty,
               SUM(amount) AS issued_value
        FROM issues
        WHERE document_date IS NOT NULL
          AND document_date != 'NaT'
          AND length(document_date) >= 7
        GROUP BY month
        ORDER BY month
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_top_suppliers() -> pd.DataFrame:
    conn = get_conn()
    query = """
        SELECT supplier,
               COUNT(DISTINCT document_no) AS total_grns,
               SUM(receive_qty) AS total_qty,
               SUM(amount) AS total_value,
               AVG(rejected_qty * 1.0 / NULLIF(receive_qty, 0)) * 100 AS rejection_pct
        FROM grn
        WHERE supplier != 'nan'
        GROUP BY supplier
        ORDER BY total_value DESC
        LIMIT 20
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def get_db_counts() -> dict:
    conn = get_conn()
    c = conn.cursor()
    counts = {}
    counts["items"]    = c.execute("SELECT COUNT(*) FROM items").fetchone()[0]
    counts["grn"]      = c.execute("SELECT COUNT(*) FROM grn").fetchone()[0]
    counts["issues"]   = c.execute("SELECT COUNT(*) FROM issues").fetchone()[0]
    counts["suppliers"]= c.execute("SELECT COUNT(DISTINCT supplier) FROM grn WHERE supplier!='nan'").fetchone()[0]
    counts["bus"]      = c.execute("SELECT COUNT(DISTINCT business_unit) FROM issues WHERE business_unit!='nan'").fetchone()[0]
    conn.close()
    return counts

# ─────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────

def main():
    init_db()
    st.set_page_config(
        page_title="SCPL Inventory",
        page_icon="📦",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.sidebar.title("📦 SCPL Inventory")
    st.sidebar.markdown("---")
    page = st.sidebar.radio("Navigation", [
        "🏠 Dashboard",
        "📥 Import Data",
        "📊 Stock Ledger",
        "🏢 Business Unit Analysis",
        "📈 Trends",
        "🚚 Supplier Performance",
        "⚠️ Low Stock Alerts",
        "⚙️ Settings"
    ])

    counts = get_db_counts()

    # ── DASHBOARD ──
    if page == "🏠 Dashboard":
        st.title("Inventory Dashboard")
        st.caption(f"SCPL — Wakad, Pune | Last updated: {datetime.now().strftime('%d %b %Y %H:%M')}")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Items",      f"{counts['items']:,}")
        c2.metric("GRN Records",      f"{counts['grn']:,}")
        c3.metric("Issue Records",    f"{counts['issues']:,}")
        c4.metric("Suppliers",        f"{counts['suppliers']:,}")
        c5.metric("Business Units",   f"{counts['bus']:,}")

        if counts["items"] > 0:
            st.markdown("---")
            df = get_stock_summary()
            low_stock = df[df["current_stock"] < df["reorder_level"]]
            zero_stock = df[df["current_stock"] <= 0]

            col1, col2, col3 = st.columns(3)
            col1.metric("Low Stock Items",  len(low_stock),  delta=f"-{len(low_stock)} need attention", delta_color="inverse")
            col2.metric("Zero Stock Items", len(zero_stock), delta_color="inverse")
            col3.metric("Total Stock Value", f"₹{df['total_grn_amount'].sum():,.0f}")

            st.markdown("---")
            col_l, col_r = st.columns(2)
            with col_l:
                st.subheader("Top 10 items by stock value")
                top10 = df.nlargest(10, "total_grn_amount")[["item_code","description","current_stock","unit","total_grn_amount"]]
                top10.columns = ["Code","Description","Stock","Unit","Value (₹)"]
                st.dataframe(top10, use_container_width=True, hide_index=True)
            with col_r:
                st.subheader("Stock by top group")
                grp = df.groupby("top_group")["current_stock"].sum().reset_index()
                grp = grp[grp["top_group"].str.len() > 2].nlargest(10, "current_stock")
                fig = px.bar(grp, x="current_stock", y="top_group", orientation="h",
                             color="current_stock", color_continuous_scale="Blues",
                             labels={"current_stock":"Stock Qty","top_group":"Group"})
                fig.update_layout(showlegend=False, coloraxis_showscale=False,
                                  margin=dict(l=0,r=0,t=10,b=0), height=320)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No data yet. Go to **Import Data** to upload your Excel files.")

    # ── IMPORT DATA ──
    elif page == "📥 Import Data":
        st.title("Import Data")

        tab1, tab2 = st.tabs(["Upload GRN Excel", "Upload Issues Excel"])

        with tab1:
            st.subheader("Upload GRN Register Excel")
            st.info("Upload the file: Goods_Receipt_Note_Register.xlsx")
            grn_file = st.file_uploader("Choose GRN Excel file", type=["xlsx"], key="grn_upload")
            if grn_file:
                with st.spinner("Parsing GRN Excel..."):
                    df = parse_grn_excel(grn_file)
                st.success(f"Parsed {len(df):,} records")
                st.dataframe(df.head(5), use_container_width=True)
                if st.button("Load GRN into Database", type="primary"):
                    with st.spinner("Loading into database... (this may take a few minutes for large files)"):
                        loaded = load_grn_to_db(df)
                    st.success(f"Loaded {loaded:,} GRN records into database!")
                    st.rerun()

        with tab2:
            st.subheader("Upload Issue Register Excel")
            st.info("Upload the file: Issue_Register.xlsx")
            iss_file = st.file_uploader("Choose Issues Excel file", type=["xlsx"], key="iss_upload")
            if iss_file:
                with st.spinner("Parsing Issues Excel..."):
                    df = parse_issues_excel(iss_file)
                st.success(f"Parsed {len(df):,} records")
                st.dataframe(df.head(5), use_container_width=True)
                if st.button("Load Issues into Database", type="primary"):
                    with st.spinner("Loading into database..."):
                        loaded = load_issues_to_db(df)
                    st.success(f"Loaded {loaded:,} issue records into database!")
                    st.rerun()

    # ── STOCK LEDGER ──
    elif page == "📊 Stock Ledger":
        st.title("Stock Ledger")
        if counts["items"] == 0:
            st.warning("No data. Please import Excel files first.")
            return
        df = get_stock_summary()
        col1, col2, col3 = st.columns(3)
        with col1:
            search = st.text_input("Search item code / description")
        with col2:
            groups = ["All"] + sorted(df["top_group"].dropna().unique().tolist())
            sel_group = st.selectbox("Filter by group", groups)
        with col3:
            show_only = st.selectbox("Show", ["All items", "Low stock only", "Zero stock only"])

        filtered = df.copy()
        if search:
            filtered = filtered[
                filtered["item_code"].str.contains(search, case=False, na=False) |
                filtered["description"].str.contains(search, case=False, na=False)
            ]
        if sel_group != "All":
            filtered = filtered[filtered["top_group"] == sel_group]
        if show_only == "Low stock only":
            filtered = filtered[filtered["current_stock"] < filtered["reorder_level"]]
        elif show_only == "Zero stock only":
            filtered = filtered[filtered["current_stock"] <= 0]

        st.caption(f"Showing {len(filtered):,} of {len(df):,} items")
        st.dataframe(
            filtered[["item_code","description","unit","top_group",
                       "total_received","total_issued","current_stock","reorder_level"]].rename(columns={
                "item_code":"Item Code","description":"Description","unit":"Unit",
                "top_group":"Group","total_received":"Total In","total_issued":"Total Out",
                "current_stock":"Current Stock","reorder_level":"Reorder Level"
            }),
            use_container_width=True, hide_index=True
        )
        csv = filtered.to_csv(index=False).encode("utf-8")
        st.download_button("Download CSV", csv, "stock_ledger.csv", "text/csv")

    # ── BUSINESS UNIT ANALYSIS ──
    elif page == "🏢 Business Unit Analysis":
        st.title("Business Unit Analysis")
        if counts["issues"] == 0:
            st.warning("No issue data. Please import Issues Excel first.")
            return
        bu_df = get_bu_summary()
        fig = px.bar(bu_df.head(15), x="total_issued_value", y="business_unit",
                     orientation="h", color="total_issued_value",
                     color_continuous_scale="Teal",
                     labels={"total_issued_value":"Issued Value (₹)","business_unit":"Business Unit"},
                     title="Top 15 Business Units by Issue Value")
        fig.update_layout(showlegend=False, coloraxis_showscale=False, height=500,
                          margin=dict(l=0,r=0,t=40,b=0))
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(bu_df.rename(columns={
            "business_unit":"Business Unit","unique_items":"Unique Items",
            "total_issued_qty":"Total Qty Issued","total_issued_value":"Total Value (₹)"
        }), use_container_width=True, hide_index=True)

    # ── TRENDS ──
    elif page == "📈 Trends":
        st.title("Monthly Trends")
        if counts["issues"] == 0:
            st.warning("No issue data. Please import Issues Excel first.")
            return
        trend = get_monthly_trend()
        if trend.empty:
            st.warning("No trend data available.")
            return
        fig = px.line(trend, x="month", y="issued_value",
                      title="Monthly Issue Value (₹)",
                      labels={"month":"Month","issued_value":"Value (₹)"})
        fig.update_traces(line_color="#378ADD", line_width=2)
        fig.update_layout(margin=dict(l=0,r=0,t=40,b=0))
        st.plotly_chart(fig, use_container_width=True)

        fig2 = px.bar(trend, x="month", y="issued_qty",
                      title="Monthly Issue Quantity",
                      labels={"month":"Month","issued_qty":"Quantity"},
                      color="issued_qty", color_continuous_scale="Blues")
        fig2.update_layout(showlegend=False, coloraxis_showscale=False,
                           margin=dict(l=0,r=0,t=40,b=0))
        st.plotly_chart(fig2, use_container_width=True)

    # ── SUPPLIER PERFORMANCE ──
    elif page == "🚚 Supplier Performance":
        st.title("Supplier Performance")
        if counts["grn"] == 0:
            st.warning("No GRN data. Please import GRN Excel first.")
            return
        sup_df = get_top_suppliers()
        sup_df["rejection_pct"] = sup_df["rejection_pct"].fillna(0).round(2)
        sup_df["total_value"] = sup_df["total_value"].round(0)

        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(sup_df.head(10), x="total_value", y="supplier",
                         orientation="h", title="Top 10 suppliers by value",
                         labels={"total_value":"Value (₹)","supplier":"Supplier"},
                         color="total_value", color_continuous_scale="Greens")
            fig.update_layout(showlegend=False, coloraxis_showscale=False,
                              margin=dict(l=0,r=0,t=40,b=0), height=350)
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            high_rej = sup_df[sup_df["rejection_pct"] > 0].nlargest(10, "rejection_pct")
            if not high_rej.empty:
                fig2 = px.bar(high_rej, x="rejection_pct", y="supplier",
                              orientation="h", title="Top 10 suppliers by rejection %",
                              labels={"rejection_pct":"Rejection %","supplier":"Supplier"},
                              color="rejection_pct", color_continuous_scale="Reds")
                fig2.update_layout(showlegend=False, coloraxis_showscale=False,
                                   margin=dict(l=0,r=0,t=40,b=0), height=350)
                st.plotly_chart(fig2, use_container_width=True)

        st.dataframe(sup_df.rename(columns={
            "supplier":"Supplier","total_grns":"GRNs","total_qty":"Total Qty",
            "total_value":"Value (₹)","rejection_pct":"Rejection %"
        }), use_container_width=True, hide_index=True)

    # ── LOW STOCK ALERTS ──
    elif page == "⚠️ Low Stock Alerts":
        st.title("Low Stock Alerts")
        if counts["items"] == 0:
            st.warning("No data. Please import Excel files first.")
            return
        df = get_stock_summary()
        low = df[df["current_stock"] < df["reorder_level"]]
        zero = df[df["current_stock"] <= 0]

        tab1, tab2 = st.tabs([f"Below Reorder Level ({len(low)})", f"Zero Stock ({len(zero)})"])
        with tab1:
            if low.empty:
                st.success("All items are above reorder level!")
            else:
                st.warning(f"{len(low)} items need attention")
                st.dataframe(low[["item_code","description","unit","current_stock","reorder_level"]].rename(columns={
                    "item_code":"Item Code","description":"Description","unit":"Unit",
                    "current_stock":"Current Stock","reorder_level":"Reorder Level"
                }), use_container_width=True, hide_index=True)
        with tab2:
            if zero.empty:
                st.success("No zero-stock items!")
            else:
                st.error(f"{len(zero)} items have zero or negative stock")
                st.dataframe(zero[["item_code","description","unit","current_stock"]].rename(columns={
                    "item_code":"Item Code","description":"Description",
                    "unit":"Unit","current_stock":"Current Stock"
                }), use_container_width=True, hide_index=True)

    # ── SETTINGS ──
    elif page == "⚙️ Settings":
        st.title("Settings")
        st.subheader("Set Reorder Levels")
        st.info("Set minimum stock level for each item. Alert triggers when stock falls below this.")
        if counts["items"] == 0:
            st.warning("No items yet. Import data first.")
            return
        conn = get_conn()
        items_df = pd.read_sql_query(
            "SELECT i.item_code, i.description, i.unit, COALESCE(r.reorder_level,0) AS reorder_level FROM items i LEFT JOIN reorder_levels r ON i.item_code=r.item_code LIMIT 100",
            conn)
        conn.close()
        search = st.text_input("Search item to set reorder level")
        if search:
            items_df = items_df[items_df["item_code"].str.contains(search, case=False, na=False) |
                                items_df["description"].str.contains(search, case=False, na=False)]
        sel = st.selectbox("Select item", items_df["item_code"] + " — " + items_df["description"])
        if sel:
            item_code = sel.split(" — ")[0]
            curr = items_df[items_df["item_code"]==item_code]["reorder_level"].values[0]
            new_level = st.number_input("Reorder level", value=float(curr), min_value=0.0)
            if st.button("Save Reorder Level"):
                conn = get_conn()
                conn.execute("INSERT OR REPLACE INTO reorder_levels (item_code, reorder_level) VALUES (?,?)",
                             (item_code, new_level))
                conn.commit()
                conn.close()
                st.success(f"Reorder level set to {new_level} for {item_code}")

        st.markdown("---")
        st.subheader("Database info")
        st.json(counts)
        if st.button("Clear all data (reset)", type="secondary"):
            conn = get_conn()
            conn.executescript("DELETE FROM grn; DELETE FROM issues; DELETE FROM items; DELETE FROM reorder_levels;")
            conn.commit()
            conn.close()
            st.success("Database cleared.")
            st.rerun()

if __name__ == "__main__":
    main()
