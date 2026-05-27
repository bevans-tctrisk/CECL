"""
TCT Credit Migration Analytics Hub - Interactive Dashboard
Streamlit dashboard with per-CU analytics, credit migration, trend analysis.

Usage: streamlit run dashboard.py
"""
import os
import sys

import streamlit as st
import pandas as pd
import numpy as np
import yaml
import plotly.express as px
import plotly.graph_objects as go
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

from cecl_engine import (
    calculate_cecl, risk_change_matrix, pool_summary,
    migration_summary_by_pool, grade_distribution, trend_data,
    build_grade_order,
)
from generate_report import load_historical_data

load_dotenv()

st.set_page_config(page_title="TCT Credit Migration Analytics Hub", layout="wide")

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))
CONFIG_FOLDER = os.path.join(BASE_FOLDER, 'client_configs')

db_url = os.getenv('DATABASE_URL')
if not db_url:
    st.error("DATABASE_URL not set. Create a .env file with DATABASE_URL=postgresql://...")
    st.stop()
engine = create_engine(db_url)


def load_client_configs():
    """Load all client configs."""
    configs = {}
    for f in os.listdir(CONFIG_FOLDER):
        if f.endswith('.yaml') and not f.startswith('_'):
            name = os.path.splitext(f)[0]
            with open(os.path.join(CONFIG_FOLDER, f), 'r', encoding='utf-8') as fh:
                configs[name] = yaml.safe_load(fh)
    return configs


@st.cache_data(ttl=3600)
def load_all_data():
    query = "SELECT * FROM monthly_loan_data"
    return pd.read_sql(query, engine)


# Header
st.title("🏦 TCT Credit Migration Analytics Hub")
col_refresh, col_spacer = st.columns([1, 5])
with col_refresh:
    if st.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.rerun()

df_raw = load_all_data()
configs = load_client_configs()

if df_raw.empty:
    st.warning("No loan data found in database. Import data first using `python import_data.py --client <name>`")
    st.stop()

# --- SIDEBAR FILTERS ---
st.sidebar.header("🔧 Filters")

cu_list = sorted(df_raw['credit_union'].dropna().unique())
selected_cu = st.sidebar.selectbox("Credit Union", cu_list, key="cu_selector")

# Find matching config
cu_config = None
cu_grades = None
for name, cfg in configs.items():
    if cfg['credit_union'] == selected_cu:
        cu_config = cfg
        cu_grades = cfg['credit_grades']
        break

# Fallback grades if no config found
if not cu_grades:
    cu_grades = [
        {"label": "A+", "min_score": 720, "max_score": 900, "reserve_rate": 0.0011},
        {"label": "A", "min_score": 680, "max_score": 719, "reserve_rate": 0.0025},
        {"label": "B", "min_score": 640, "max_score": 679, "reserve_rate": 0.0050},
        {"label": "C", "min_score": 620, "max_score": 639, "reserve_rate": 0.0116},
        {"label": "D", "min_score": 600, "max_score": 619, "reserve_rate": 0.0250},
        {"label": "E", "min_score": 0, "max_score": 599, "reserve_rate": 0.0500},
    ]

no_score_label = cu_config.get('no_score_label', 'Not Reported') if cu_config else 'Not Reported'

# Date filter
df_cu_raw = df_raw[df_raw['credit_union'] == selected_cu]
date_list = sorted(df_cu_raw['snapshot_date'].dropna().unique(), reverse=True)

if not date_list:
    st.sidebar.warning("No dates found for this Credit Union.")
    st.stop()

selected_date = st.sidebar.selectbox("Quarter Ending", date_list, key="date_selector")

# Pool filter
df_period = df_cu_raw[df_cu_raw['snapshot_date'] == selected_date]
pool_list = sorted(df_period['loan_pool'].dropna().unique())
selected_pools = st.sidebar.multiselect("Loan Pools", pool_list, default=pool_list, key="pool_selector")

# Apply CECL calculations
df_calc = calculate_cecl(df_period, cu_grades, no_score_label)
df = df_calc[df_calc['loan_pool'].isin(selected_pools)]

# Calculate all-period data for trends
df_all_calc = calculate_cecl(df_cu_raw, cu_grades, no_score_label)

# Load historical file-based data if configured
hist = None
if cu_config and cu_config.get('data_directory'):
    hist = load_historical_data(cu_config)

# ===== TAB LAYOUT =====
tab_overview, tab_migration, tab_pools, tab_grades, tab_trends, tab_hist, tab_detail = st.tabs([
    "📊 Overview", "🔄 Credit Migration", "🏷️ Loan Pools",
    "📈 Grade Distribution", "📉 Trends", "📂 Historical Data", "📋 Loan Detail"
])

# ===== TAB 1: OVERVIEW =====
with tab_overview:
    st.header(f"{selected_cu}")
    st.subheader(f"Quarter Ending: {selected_date}")
    st.divider()

    total_bal = df['current_balance'].sum()
    total_acl = df['expected_loss_amount'].sum()
    loan_count = len(df)
    reserve_pct = total_acl / total_bal if total_bal > 0 else 0
    improved = (df['migration_status'] == 'Improved').sum()
    deteriorated = (df['migration_status'] == 'Deteriorated').sum()
    unchanged = (df['migration_status'] == 'Unchanged').sum()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Portfolio Balance", f"${total_bal:,.2f}")
    m2.metric("Total Loans", f"{loan_count:,}")
    m3.metric("Required ACL Reserve", f"${total_acl:,.2f}")
    m4.metric("Reserve %", f"{reserve_pct:.4%}")

    m5, m6, m7, m8 = st.columns(4)
    m5.metric("Improved", f"{improved:,}", delta=f"{improved}", delta_color="normal")
    m6.metric("Deteriorated", f"{deteriorated:,}", delta=f"-{deteriorated}" if deteriorated > 0 else "0",
              delta_color="normal")
    m7.metric("Unchanged", f"{unchanged:,}")
    m8.metric("Avg FICO", f"{df[df['current_fico_score'] > 0]['current_fico_score'].mean():.0f}"
              if len(df[df['current_fico_score'] > 0]) > 0 else "N/A")

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Reserve by Loan Pool")
        ps = pool_summary(df)
        fig_bar = px.bar(ps, x='loan_pool', y='total_reserve', color='loan_pool',
                         labels={'total_reserve': 'Reserve Amount', 'loan_pool': 'Loan Pool'},
                         text_auto='$.2s')
        fig_bar.update_layout(showlegend=False)
        st.plotly_chart(fig_bar, use_container_width=True)

    with c2:
        st.subheader("Credit Migration Status")
        mig_data = pd.DataFrame({
            'Status': ['Improved', 'Deteriorated', 'Unchanged'],
            'Count': [improved, deteriorated, unchanged]
        })
        fig_pie = px.pie(mig_data, names='Status', values='Count', hole=0.4,
                         color='Status',
                         color_discrete_map={
                             'Improved': '#2ecc71',
                             'Deteriorated': '#e74c3c',
                             'Unchanged': '#95a5a6'
                         })
        st.plotly_chart(fig_pie, use_container_width=True)


# ===== TAB 2: CREDIT MIGRATION =====
with tab_migration:
    st.header("Risk Change by Credit Score")
    st.caption("Rows = Current Grade, Columns = Original Grade. Green = Improved, Red = Deteriorated.")

    matrix = risk_change_matrix(df, cu_grades, no_score_label)

    # Format as heatmap
    grade_labels = [g['label'] for g in cu_grades] + [no_score_label]
    matrix_display = matrix.loc[
        [gl for gl in grade_labels if gl in matrix.index],
        [gl for gl in grade_labels if gl in matrix.columns]
    ]

    fig_heat = go.Figure(data=go.Heatmap(
        z=matrix_display.values,
        x=matrix_display.columns.tolist(),
        y=matrix_display.index.tolist(),
        colorscale='RdYlGn_r',
        text=[[f"${v:,.0f}" for v in row] for row in matrix_display.values],
        texttemplate="%{text}",
        hovertemplate="Current: %{y}<br>Original: %{x}<br>Balance: %{text}<extra></extra>",
    ))
    fig_heat.update_layout(
        xaxis_title="Original Credit Grade",
        yaxis_title="Current Credit Grade",
        yaxis_autorange='reversed',
        height=500,
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    # Migration summary table (formatted to match TCT report style)
    st.subheader("Migration Dollar Summary")

    range_overrides = cu_config.get('risk_change_range_labels', {}) if cu_config else {}
    grade_ranges = {g['label']: range_overrides.get(g['label'], f"{g['min_score']}-{g['max_score']}") for g in cu_grades}
    grade_ranges[no_score_label] = ""

    row_labels = [gl for gl in grade_labels if gl in matrix_display.index]
    col_labels = [gl for gl in grade_labels if gl in matrix_display.columns]

    table_df = pd.DataFrame({
        'Current Credit Grade': row_labels,
        'Range': [grade_ranges.get(gl, '') for gl in row_labels],
    })

    for col in col_labels:
        table_df[col] = [matrix_display.loc[row, col] for row in row_labels]

    table_df['Grand Total'] = table_df[col_labels].sum(axis=1)

    total_row = {'Current Credit Grade': 'Total', 'Range': ''}
    for col in col_labels:
        total_row[col] = table_df[col].sum()
    total_row['Grand Total'] = table_df['Grand Total'].sum()
    table_df = pd.concat([table_df, pd.DataFrame([total_row])], ignore_index=True)

    st.markdown("**Dollar**")

    def fmt_currency(val):
        try:
            v = float(val)
            return '-' if abs(v) < 0.005 else f"${v:,.0f}"
        except Exception:
            return val

    grade_only = [g['label'] for g in cu_grades]
    grade_pos = {g: i for i, g in enumerate(grade_only)}

    def style_matrix_cells(row):
        styles = [''] * len(row)
        row_label = row['Current Credit Grade']
        for idx, col_name in enumerate(row.index):
            if col_name in col_labels and row_label in grade_pos and col_name in grade_pos:
                if grade_pos[col_name] > grade_pos[row_label]:
                    styles[idx] = 'background-color: #dbeac1;'
                elif grade_pos[col_name] < grade_pos[row_label]:
                    styles[idx] = 'background-color: #e8c3c3;'
        return styles

    display_columns = ['Current Credit Grade', 'Range'] + col_labels + ['Grand Total']
    num_cols = col_labels + ['Grand Total']

    styled = (
        table_df[display_columns]
        .style
        .format({c: fmt_currency for c in num_cols})
        .apply(style_matrix_cells, axis=1)
        .set_properties(**{'text-align': 'center'})
        .set_properties(subset=['Current Credit Grade', 'Range'], **{'text-align': 'left', 'font-weight': 'bold'})
        .set_table_styles([
            {'selector': 'th', 'props': [('background-color', '#d9d9d9'), ('color', '#000000'), ('font-weight', 'bold'), ('text-align', 'center')]},
            {'selector': 'td', 'props': [('border', '1px solid #666666')]},
            {'selector': 'th.col_heading', 'props': [('border', '1px solid #666666')]},
            {'selector': 'th.row_heading', 'props': [('border', '1px solid #666666')]},
        ])
    )
    st.dataframe(styled, use_container_width=True, height=360)


# ===== TAB 3: LOAN POOLS =====
with tab_pools:
    st.header("Loan Pool Analysis")

    ps = pool_summary(df)
    st.dataframe(
        ps.style.format({
            'total_balance': '${:,.2f}',
            'total_reserve': '${:,.2f}',
            'reserve_pct': '{:.4%}',
        }),
        use_container_width=True,
    )

    # Pool breakdown charts
    c1, c2 = st.columns(2)
    with c1:
        fig_pool_bal = px.treemap(ps, path=['loan_pool'], values='total_balance',
                                  color='reserve_pct', color_continuous_scale='RdYlGn_r',
                                  title="Balance by Pool (colored by reserve rate)")
        st.plotly_chart(fig_pool_bal, use_container_width=True)

    with c2:
        # Migration by pool
        mig_pool = migration_summary_by_pool(df)
        fig_mig = px.bar(mig_pool, x='loan_pool', y='balance', color='migration_status',
                         barmode='group',
                         color_discrete_map={
                             'Improved': '#2ecc71',
                             'Deteriorated': '#e74c3c',
                             'Unchanged': '#95a5a6'
                         },
                         title="Migration Status by Pool")
        st.plotly_chart(fig_mig, use_container_width=True)


# ===== TAB 4: GRADE DISTRIBUTION =====
with tab_grades:
    st.header("Credit Grade Distribution")

    gd = grade_distribution(df)
    grade_order = [g['label'] for g in cu_grades] + [no_score_label]
    gd['sort_order'] = gd['current_grade'].apply(
        lambda x: grade_order.index(x) if x in grade_order else 99
    )
    gd = gd.sort_values('sort_order').drop(columns='sort_order')

    st.dataframe(
        gd.style.format({
            'total_balance': '${:,.2f}',
            'total_reserve': '${:,.2f}',
            'reserve_pct': '{:.4%}',
        }),
        use_container_width=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        fig_grade_bar = px.bar(gd, x='current_grade', y='total_balance',
                                color='current_grade', title="Balance by Grade",
                                text_auto='$.2s')
        fig_grade_bar.update_layout(showlegend=False)
        st.plotly_chart(fig_grade_bar, use_container_width=True)

    with c2:
        fig_grade_pie = px.pie(gd, names='current_grade', values='loan_count',
                                title="Loan Count by Grade")
        st.plotly_chart(fig_grade_pie, use_container_width=True)

    # FICO distribution histogram
    st.subheader("FICO Score Distribution")
    fico_scores = df[df['current_fico_score'] > 0]['current_fico_score']
    if len(fico_scores) > 0:
        fig_fico = px.histogram(fico_scores, nbins=30,
                                 labels={'value': 'FICO Score', 'count': 'Number of Loans'},
                                 color_discrete_sequence=['#2E86C1'])
        # Add grade boundary lines
        for g in cu_grades:
            fig_fico.add_vline(x=g['min_score'], line_dash="dash",
                                line_color="gray", annotation_text=g['label'])
        st.plotly_chart(fig_fico, use_container_width=True)


# ===== TAB 5: TRENDS =====
with tab_trends:
    st.header("Historical Trends")

    td = trend_data(df_all_calc, selected_cu)
    if len(td) > 1:
        c1, c2 = st.columns(2)
        with c1:
            fig_bal = px.line(td, x='snapshot_date', y='total_balance',
                              title="Portfolio Balance Over Time", markers=True)
            fig_bal.update_layout(yaxis_tickformat='$,.0f')
            st.plotly_chart(fig_bal, use_container_width=True)

        with c2:
            fig_res = px.line(td, x='snapshot_date', y='reserve_pct',
                              title="Reserve Rate Over Time", markers=True)
            fig_res.update_layout(yaxis_tickformat='.4%')
            st.plotly_chart(fig_res, use_container_width=True)

        # Migration trends
        st.subheader("Credit Migration Trends")
        mig_cols = ['improved_count', 'deteriorated_count', 'unchanged_count']
        if all(c in td.columns for c in mig_cols):
            td_melt = td.melt(id_vars='snapshot_date', value_vars=mig_cols,
                               var_name='Status', value_name='Count')
            td_melt['Status'] = td_melt['Status'].str.replace('_count', '').str.title()
            fig_mig_trend = px.bar(td_melt, x='snapshot_date', y='Count', color='Status',
                                    barmode='group',
                                    color_discrete_map={
                                        'Improved': '#2ecc71',
                                        'Deteriorated': '#e74c3c',
                                        'Unchanged': '#95a5a6'
                                    })
            st.plotly_chart(fig_mig_trend, use_container_width=True)

        # Trend table
        st.subheader("Quarterly Data")
        st.dataframe(
            td.style.format({
                'total_balance': '${:,.2f}',
                'total_reserve': '${:,.2f}',
                'reserve_pct': '{:.4%}',
            }),
            use_container_width=True,
        )
    else:
        st.info("Import multiple quarters to see trend analysis. Currently only 1 period loaded.")


# ===== TAB 6: HISTORICAL DATA =====
with tab_hist:
    st.header("Historical Data")

    if hist and hist.get('years'):
        years = hist['years']
        co_data = hist['chargeoffs']
        rc_data = hist['recoveries']
        avg_bals = hist.get('avg_balances', {})
        dq_pct = hist.get('dq_pct', {})
        monthly = hist.get('monthly_balances', pd.DataFrame())
        pools = sorted(set(cu_config.get('pool_map', {}).values())) if cu_config else []

        # ── Charge-off & Recovery Trends ──
        st.subheader("Charge-off & Recovery Trends")
        co_rc_rows = []
        for y in years:
            total_co = sum(co_data.get(y, {}).values())
            total_rc = sum(rc_data.get(y, {}).values())
            co_rc_rows.append({'Year': y, 'Charge-offs': total_co, 'Recoveries': total_rc,
                               'Net Charge-offs': total_co - total_rc})
        co_rc_df = pd.DataFrame(co_rc_rows)

        c1, c2 = st.columns(2)
        with c1:
            fig_co = px.bar(co_rc_df, x='Year', y=['Charge-offs', 'Recoveries'],
                            barmode='group', title="Annual Charge-offs vs Recoveries",
                            color_discrete_sequence=['#e74c3c', '#2ecc71'])
            fig_co.update_layout(yaxis_tickformat='$,.0f')
            st.plotly_chart(fig_co, use_container_width=True)

        with c2:
            fig_net = px.bar(co_rc_df, x='Year', y='Net Charge-offs',
                             title="Net Charge-offs by Year",
                             color_discrete_sequence=['#c0392b'])
            fig_net.update_layout(yaxis_tickformat='$,.0f')
            st.plotly_chart(fig_net, use_container_width=True)

        # Charge-off by pool
        st.subheader("Charge-offs by Pool")
        co_pool_rows = []
        for y in years:
            for pool in pools:
                co_pool_rows.append({'Year': y, 'Pool': pool,
                                     'Amount': co_data.get(y, {}).get(pool, 0)})
        co_pool_df = pd.DataFrame(co_pool_rows)
        fig_co_pool = px.bar(co_pool_df, x='Year', y='Amount', color='Pool',
                              title="Charge-offs by Pool by Year")
        fig_co_pool.update_layout(yaxis_tickformat='$,.0f')
        st.plotly_chart(fig_co_pool, use_container_width=True)

        # Charge-off/Recovery data table
        with st.expander("Charge-off & Recovery Detail Table"):
            detail_rows = []
            for y in years:
                for pool in pools:
                    co_amt = co_data.get(y, {}).get(pool, 0)
                    rc_amt = rc_data.get(y, {}).get(pool, 0)
                    if co_amt > 0 or rc_amt > 0:
                        detail_rows.append({'Year': y, 'Pool': pool, 'Charge-offs': co_amt,
                                            'Recoveries': rc_amt, 'Net': co_amt - rc_amt})
            if detail_rows:
                detail_df = pd.DataFrame(detail_rows)
                st.dataframe(detail_df.style.format({
                    'Charge-offs': '${:,.2f}', 'Recoveries': '${:,.2f}', 'Net': '${:,.2f}'
                }), use_container_width=True)

        st.divider()

        # ── Life Loss Rates ──
        st.subheader("Life Loss Rates by Pool")
        llr_rows = []
        for pool in pools:
            rates = []
            for y in years:
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                rates.append(rate)
                llr_rows.append({'Year': y, 'Pool': pool, 'Life Loss Rate': rate})
        llr_df = pd.DataFrame(llr_rows)

        if not llr_df.empty:
            fig_llr = px.line(llr_df, x='Year', y='Life Loss Rate', color='Pool',
                              markers=True, title="Life Loss Rate by Pool Over Time")
            fig_llr.update_layout(yaxis_tickformat='.4%')
            st.plotly_chart(fig_llr, use_container_width=True)

            # Summary table: average life loss per pool
            avg_llr = llr_df.groupby('Pool')['Life Loss Rate'].mean().reset_index()
            avg_llr.columns = ['Pool', 'Avg Life Loss Rate']
            st.dataframe(avg_llr.style.format({'Avg Life Loss Rate': '{:.4%}'}),
                         use_container_width=True)

        st.divider()

        # ── Delinquency ──
        dq_data_raw = hist.get('delinquency', {})
        if dq_data_raw:
            st.subheader("Delinquency Trends")

            dq_rows = []
            for qlabel in sorted(dq_data_raw.keys()):
                for pool, bal in dq_data_raw[qlabel].items():
                    dq_rows.append({'Quarter': qlabel, 'Pool': pool, 'DQ Balance': bal})
            dq_df = pd.DataFrame(dq_rows)

            c1, c2 = st.columns(2)
            with c1:
                dq_total = dq_df.groupby('Quarter')['DQ Balance'].sum().reset_index()
                fig_dq = px.bar(dq_total, x='Quarter', y='DQ Balance',
                                title="Total Delinquent Balance by Quarter",
                                color_discrete_sequence=['#e67e22'])
                fig_dq.update_layout(yaxis_tickformat='$,.0f')
                st.plotly_chart(fig_dq, use_container_width=True)

            with c2:
                fig_dq_pool = px.bar(dq_df, x='Quarter', y='DQ Balance', color='Pool',
                                     title="Delinquent Balance by Pool by Quarter")
                fig_dq_pool.update_layout(yaxis_tickformat='$,.0f')
                st.plotly_chart(fig_dq_pool, use_container_width=True)

            # DQ % table
            if dq_pct:
                with st.expander("Delinquency Rate (%) by Year"):
                    dq_pct_rows = []
                    for y in sorted(dq_pct.keys()):
                        row = {'Year': y}
                        for pool in pools:
                            row[pool] = dq_pct[y].get(pool, 0)
                        dq_pct_rows.append(row)
                    dq_pct_df = pd.DataFrame(dq_pct_rows)
                    fmt = {p: '{:.4%}' for p in pools}
                    st.dataframe(dq_pct_df.style.format(fmt), use_container_width=True)

            st.divider()

        # ── Historical Loan Balances ──
        if not monthly.empty:
            st.subheader("Historical Loan Balances")

            monthly_agg = monthly.groupby(['date', 'pool'])['balance'].sum().reset_index()
            fig_bal = px.area(monthly_agg, x='date', y='balance', color='pool',
                              title="Loan Balances by Pool Over Time")
            fig_bal.update_layout(yaxis_tickformat='$,.0f', xaxis_title='Date',
                                  yaxis_title='Balance')
            st.plotly_chart(fig_bal, use_container_width=True)

            # Total balance over time
            total_monthly = monthly.groupby('date')['balance'].sum().reset_index()
            fig_total = px.line(total_monthly, x='date', y='balance',
                                title="Total Portfolio Balance Over Time", markers=True)
            fig_total.update_layout(yaxis_tickformat='$,.0f')
            st.plotly_chart(fig_total, use_container_width=True)

            # Latest quarter breakdown
            with st.expander("Latest Balance Breakdown by Pool"):
                latest_date = monthly['date'].max()
                latest = monthly[monthly['date'] == latest_date].groupby('pool')['balance'].sum().reset_index()
                latest.columns = ['Pool', 'Balance']
                latest = latest.sort_values('Balance', ascending=False)
                st.dataframe(latest.style.format({'Balance': '${:,.2f}'}),
                             use_container_width=True)
    else:
        st.info("No historical data available for this credit union. "
                "Configure `data_directory` in the client YAML and add charge-off, "
                "recovery, delinquency, and balance files to enable this tab.")


# ===== TAB 7: LOAN DETAIL =====
with tab_detail:
    st.header("Detailed Loan View")

    # Filters within detail tab
    dc1, dc2 = st.columns(2)
    with dc1:
        detail_grade = st.multiselect("Filter by Grade", df['current_grade'].unique(),
                                       default=list(df['current_grade'].unique()),
                                       key="detail_grade")
    with dc2:
        detail_status = st.multiselect("Filter by Migration Status",
                                        df['migration_status'].unique(),
                                        default=list(df['migration_status'].unique()),
                                        key="detail_status")

    df_detail = df[
        (df['current_grade'].isin(detail_grade)) &
        (df['migration_status'].isin(detail_status))
    ]

    display_cols = ['member_number', 'loan_pool', 'current_balance',
                    'current_fico_score', 'current_grade', 'original_fico_score',
                    'original_grade', 'migration_status', 'reserve_rate',
                    'expected_loss_amount']

    st.dataframe(
        df_detail[display_cols].style.format({
            'current_balance': '${:,.2f}',
            'expected_loss_amount': '${:,.2f}',
            'reserve_rate': '{:.4%}',
        }),
        use_container_width=True,
        height=600,
    )

    st.caption(f"Showing {len(df_detail):,} of {len(df):,} loans")

    # Download button
    csv = df_detail[display_cols].to_csv(index=False)
    st.download_button("📥 Download Filtered Data (CSV)", csv,
                        f"{selected_cu}_{selected_date}_loans.csv",
                        "text/csv")