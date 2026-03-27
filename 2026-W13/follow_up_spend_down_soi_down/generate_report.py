"""Generate P1 Assess report: SOI Down + Spend Down accounts — styled HTML."""
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os
import re
import base64
from io import BytesIO

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(OUTPUT_DIR, sorted([f for f in os.listdir(OUTPUT_DIR) if f.endswith('.csv')])[0])

PERIOD_ORDER = ["1_Benchmark", "2_Jan", "3_Feb", "4_Mar_MTD"]
PERIOD_LABELS_CHART = {"1_Benchmark": "Benchmark\n(Dec 25-31)", "2_Jan": "Jan", "3_Feb": "Feb", "4_Mar_MTD": "Mar MTD"}
PERIOD_LABELS = {"1_Benchmark": "Benchmark", "2_Jan": "Jan", "3_Feb": "Feb", "4_Mar_MTD": "Mar MTD"}

ACCOUNT_SHORT = {
    "Last Z_IOS": "Last Z (iOS)",
    "Lands of Jail_Madhouse_04": "Lands of Jail (Madhouse)",
    "Dark War Survival_MADHOUS": "Dark War Survival (Madhouse)",
}
CJK_APP_NAMES = {}

def sanitize_app_name(name):
    return CJK_APP_NAMES.get(name, name)

def safe_filename(name):
    short = ACCOUNT_SHORT.get(name, name)
    return re.sub(r"[^\w\-]", "_", short)[:50]

def fmt_spend(v):
    if abs(v) >= 1_000_000: return f"${v/1_000_000:.1f}M"
    elif abs(v) >= 1_000: return f"${v/1_000:.1f}K"
    return f"${v:.0f}"

def fmt_spend_table(v):
    return f"${v:,.0f}"

def fmt_pct(v):
    return f"{v:.2f}%"

def fmt_int(v):
    return f"{int(v):,}"

def calc_soi(mol, non_mol):
    total = mol + non_mol
    return (mol / total * 100) if total > 0 else 0.0

def fig_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


# -- Charts ------------------------------------------------------------------

def make_account_chart_b64(df, title):
    fig, ax1 = plt.subplots(figsize=(7, 4))
    periods = [p for p in PERIOD_ORDER if p in df["period"].values]
    sub = df.set_index("period").loc[periods]
    labels = [PERIOD_LABELS_CHART.get(p, p) for p in periods]
    x = range(len(periods))

    bars = ax1.bar(x, sub["avg_daily_spend"], color="#4C9AFF", alpha=0.7, width=0.45)
    ax1.set_ylabel("Avg Daily Spend", fontsize=9, color="#4C9AFF")
    ax1.tick_params(axis="y", labelcolor="#4C9AFF", labelsize=8)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: fmt_spend(v)))
    for i, bar in enumerate(bars):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
                 fmt_spend(sub["avg_daily_spend"].iloc[i]),
                 ha="center", va="bottom", fontsize=8, color="#4C9AFF")

    ax2 = ax1.twinx()
    ax2.plot(x, sub["soi_pct"], color="#FF5252", marker="o", linewidth=2.5, markersize=8, zorder=5)
    ax2.set_ylabel("SOI (%)", fontsize=9, color="#FF5252")
    ax2.tick_params(axis="y", labelcolor="#FF5252", labelsize=8)
    for i, v in enumerate(sub["soi_pct"]):
        ax2.annotate(f"{v:.1f}%", (i, v), textcoords="offset points", xytext=(0, 10),
                     ha="center", fontsize=9, fontweight="bold", color="#FF5252")

    ax1.set_xticks(x); ax1.set_xticklabels(labels, fontsize=9)
    ax1.set_title(title, fontsize=11, fontweight="bold", pad=10)
    ax1.set_ylim(bottom=0)
    soi_vals = sub["soi_pct"]
    margin = max((soi_vals.max() - soi_vals.min()) * 0.5, 1.0)
    ax2.set_ylim(max(0, soi_vals.min() - margin), soi_vals.max() + margin * 2)
    ax1.grid(axis="y", alpha=0.2)
    plt.tight_layout()
    return fig_to_base64(fig)


def make_app_charts_b64(df_apps, account_name):
    apps_total = df_apps.groupby("app_name")["avg_daily_spend"].sum().sort_values(ascending=False)
    top_apps = apps_total.head(8).index.tolist()
    df_top = df_apps[df_apps["app_name"].isin(top_apps)].copy()
    if df_top.empty:
        return None

    n = len(top_apps)
    ncols = min(4, n)
    nrows = (n + ncols - 1) // ncols
    fig, axes = plt.subplots(nrows, ncols, figsize=(4.2 * ncols, 3.5 * nrows))
    if nrows == 1 and ncols == 1: axes = np.array([[axes]])
    elif nrows == 1: axes = axes.reshape(1, -1)
    elif ncols == 1: axes = axes.reshape(-1, 1)

    for idx, app in enumerate(top_apps):
        row, col = divmod(idx, ncols)
        ax1 = axes[row][col]
        grp = df_top[df_top["app_name"] == app]
        periods = [p for p in PERIOD_ORDER if p in grp["period"].values]
        if not periods: ax1.set_visible(False); continue
        sub = grp.set_index("period").loc[periods]
        x = range(len(periods))

        ax1.bar(x, sub["avg_daily_spend"], color="#4C9AFF", alpha=0.6, width=0.45)
        ax1.set_ylabel("DRR", fontsize=7, color="#4C9AFF")
        ax1.tick_params(axis="y", labelcolor="#4C9AFF", labelsize=6)
        ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: fmt_spend(v)))
        ax1.set_ylim(bottom=0)

        ax2 = ax1.twinx()
        ax2.plot(x, sub["soi_pct"], color="#FF5252", marker="o", linewidth=2, markersize=6, zorder=5)
        ax2.set_ylabel("SOI%", fontsize=7, color="#FF5252")
        ax2.tick_params(axis="y", labelcolor="#FF5252", labelsize=6)
        for i, v in enumerate(sub["soi_pct"]):
            ax2.annotate(f"{v:.1f}%", (i, v), textcoords="offset points", xytext=(0, 8),
                         ha="center", fontsize=7, fontweight="bold", color="#FF5252")
        soi_vals = sub["soi_pct"]
        margin = max((soi_vals.max() - soi_vals.min()) * 0.5, 1.0)
        ax2.set_ylim(max(0, soi_vals.min() - margin), soi_vals.max() + margin * 2)

        ax1.set_xticks(x)
        ax1.set_xticklabels([PERIOD_LABELS_CHART.get(p, p) for p in periods], fontsize=7)
        ax1.set_title(sanitize_app_name(app)[:35], fontsize=8, fontweight="bold")
        ax1.grid(axis="y", alpha=0.15)

    for idx in range(n, nrows * ncols):
        r, c = divmod(idx, ncols); axes[r][c].set_visible(False)

    short = ACCOUNT_SHORT.get(account_name, account_name)
    fig.suptitle(f"{short} — App SOI vs Spend", fontsize=11, fontweight="bold", y=1.01)
    plt.tight_layout()
    return fig_to_base64(fig)


# -- HTML builders ------------------------------------------------------------

def html_account_table(df_total):
    rows_html = []
    for p in PERIOD_ORDER:
        sub = df_total[df_total["period"] == p]
        if sub.empty: continue
        r = sub.iloc[0]
        rows_html.append(f"""<tr>
            <td>{PERIOD_LABELS[p]}</td>
            <td class="num">{fmt_spend_table(r['total_spend'])}</td>
            <td class="num">{fmt_spend_table(r['avg_daily_spend'])}</td>
            <td class="num">{fmt_int(r['moloco_installs'])}</td>
            <td class="num">{fmt_int(r['non_moloco_installs'])}</td>
            <td class="num">{fmt_pct(r['soi_pct'])}</td>
        </tr>""")
    return f"""<table>
        <thead><tr>
            <th>Period</th><th class="num">Total Spend</th><th class="num">Avg Daily Spend</th>
            <th class="num">Moloco Inst</th><th class="num">Non-Moloco Inst</th><th class="num">SOI</th>
        </tr></thead>
        <tbody>{"".join(rows_html)}</tbody>
    </table>"""


def html_impact_table(df_apps, df_total):
    bm = df_apps[df_apps["period"] == "1_Benchmark"].copy()
    mar = df_apps[df_apps["period"] == "4_Mar_MTD"].copy()
    bm_t = df_total[df_total["period"] == "1_Benchmark"]
    mar_t = df_total[df_total["period"] == "4_Mar_MTD"]
    if bm_t.empty or mar_t.empty: return ""

    bm_denom = bm_t.iloc[0]["moloco_installs"] + bm_t.iloc[0]["non_moloco_installs"]
    mar_denom = mar_t.iloc[0]["moloco_installs"] + mar_t.iloc[0]["non_moloco_installs"]
    if bm_denom == 0 or mar_denom == 0: return ""

    mar_total_spend = mar["avg_daily_spend"].sum()
    impacts = []
    all_apps = set(bm["app_name"].tolist() + mar["app_name"].tolist())
    for app in all_apps:
        bm_mol = bm.loc[bm["app_name"] == app, "moloco_installs"].sum()
        mar_mol = mar.loc[mar["app_name"] == app, "moloco_installs"].sum()
        bm_nonmol = bm.loc[bm["app_name"] == app, "non_moloco_installs"].sum()
        mar_nonmol = mar.loc[mar["app_name"] == app, "non_moloco_installs"].sum()
        bm_soi = calc_soi(bm_mol, bm_nonmol)
        mar_soi = calc_soi(mar_mol, mar_nonmol)
        mar_spend = mar.loc[mar["app_name"] == app, "avg_daily_spend"].sum()
        spend_share = (mar_spend / mar_total_spend * 100) if mar_total_spend > 0 else 0
        bm_c = (bm_mol / bm_denom * 100) if bm_denom > 0 else 0
        mar_c = (mar_mol / mar_denom * 100) if mar_denom > 0 else 0
        impacts.append({"app": app, "bm_soi": bm_soi, "mar_soi": mar_soi,
                        "soi_chg": mar_soi - bm_soi, "mar_spend": mar_spend,
                        "spend_share": spend_share, "impact": mar_c - bm_c})

    impacts.sort(key=lambda x: x["impact"], reverse=True)

    rows = []
    for r in impacts:
        name = sanitize_app_name(r["app"])[:35]
        chg_cls = "positive" if r["soi_chg"] >= 0 else "negative"
        imp_cls = "positive" if r["impact"] >= 0 else "negative"
        chg_s = f"+{r['soi_chg']:.2f}%" if r["soi_chg"] >= 0 else f"{r['soi_chg']:.2f}%"
        imp_s = f"+{r['impact']:.3f}" if r["impact"] >= 0 else f"{r['impact']:.3f}"
        is_top = r["spend_share"] >= 10
        name_html = f"<strong>{name}</strong>" if is_top else name
        row_cls = ' class="highlight"' if is_top else ""
        rows.append(f"""<tr{row_cls}>
            <td>{name_html}</td>
            <td class="num">{"—" if r['bm_soi'] == 0 else fmt_pct(r['bm_soi'])}</td>
            <td class="num">{"—" if r['mar_soi'] == 0 else fmt_pct(r['mar_soi'])}</td>
            <td class="num {chg_cls}">{chg_s}</td>
            <td class="num">{"—" if r['mar_spend'] == 0 else fmt_spend_table(r['mar_spend'])}</td>
            <td class="num">{"—" if r['spend_share'] == 0 else f"{r['spend_share']:.1f}%"}</td>
            <td class="num {imp_cls}"><strong>{imp_s}</strong></td>
        </tr>""")

    return f"""<p class="note">Which apps are dragging SOI down as spend declines? <strong>SOI Impact</strong> = change in app's Moloco install share of total installs. Negative = Moloco installs dropping faster than overall market. <span class="highlight-label">Highlighted rows</span> = top spend apps (&ge;10% share).</p>
    <table>
        <thead><tr>
            <th>App</th><th class="num">BM SOI</th><th class="num">Mar SOI</th>
            <th class="num">SOI Chg</th><th class="num">Mar DRR</th>
            <th class="num">Spend %</th><th class="num">Impact (pp)</th>
        </tr></thead>
        <tbody>{"".join(rows)}</tbody>
    </table>"""


def html_app_pivot(df_apps):
    apps_by_spend = df_apps.groupby("app_name")["avg_daily_spend"].sum().sort_values(ascending=False)

    header = "<th>App</th>"
    for p in PERIOD_ORDER:
        lbl = PERIOD_LABELS[p]
        header += f'<th class="num">{lbl} DRR</th><th class="num">{lbl} SOI</th>'

    rows = []
    for app_name in apps_by_spend.index:
        app_data = df_apps[df_apps["app_name"] == app_name].set_index("period")
        name = sanitize_app_name(app_name)[:35]
        cells = f"<td><strong>{name}</strong></td>"
        for p in PERIOD_ORDER:
            if p in app_data.index:
                r = app_data.loc[p]
                cells += f'<td class="num">{fmt_spend_table(r["avg_daily_spend"])}</td>'
                cells += f'<td class="num">{fmt_pct(r["soi_pct"])}</td>'
            else:
                cells += '<td class="num dim">—</td><td class="num dim">—</td>'
        rows.append(f"<tr>{cells}</tr>")

    return f"""<div class="table-scroll"><table>
        <thead><tr>{header}</tr></thead>
        <tbody>{"".join(rows)}</tbody>
    </table></div>"""


# -- Main ---------------------------------------------------------------------

CSS = """
<style>
    * { box-sizing: border-box; }
    @page { size: landscape; margin: 8mm; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
           max-width: 100%; margin: 0; padding: 10px 15px; color: #1a1a1a;
           background: #fff; line-height: 1.4; font-size: 11px; }
    h1 { border-bottom: 2px solid #4C9AFF; padding-bottom: 6px; font-size: 18px; margin: 8px 0; }
    h2 { margin-top: 20px; padding: 5px 10px; background: #e8f0fe; border-left: 3px solid #4C9AFF;
         font-size: 14px; border-radius: 0 3px 3px 0; margin-bottom: 8px; }
    h3 { font-size: 12px; color: #444; margin: 12px 0 6px 0; }
    table { border-collapse: collapse; width: 100%; margin: 6px 0 10px 0; font-size: 10px; }
    th { background: #f0f3f8; padding: 4px 6px; text-align: left; border-bottom: 2px solid #ccc;
         font-weight: 600; white-space: nowrap; }
    td { padding: 3px 6px; border-bottom: 1px solid #e8e8e8; white-space: nowrap; }
    th.num, td.num { text-align: right; font-variant-numeric: tabular-nums; }
    tr:hover { background: #f5f8ff; }
    tr.highlight { background: #fffde7; }
    tr.highlight:hover { background: #fff9c4; }
    .positive { color: #2e7d32; }
    .negative { color: #c62828; }
    .dim { color: #bbb; }
    .note { font-size: 10px; color: #555; background: #f5f5f5; padding: 5px 8px;
            border-left: 3px solid #4C9AFF; border-radius: 0 3px 3px 0; margin: 4px 0; }
    .highlight-label { background: #fffde7; padding: 1px 4px; border: 1px solid #e0d87a; border-radius: 3px; font-size: 9px; }
    .chart { text-align: center; margin: 8px 0; }
    .chart img { max-width: 100%; height: auto; border: 1px solid #e0e0e0; border-radius: 3px; }
    .table-scroll { overflow-x: visible; }
    .meta { font-size: 10px; color: #666; }
    .toc { background: #fff; border: 1px solid #e0e0e0; border-radius: 4px; padding: 8px 14px; display: inline-block; font-size: 11px; }
    .toc a { text-decoration: none; color: #1565c0; }
    .toc a:hover { text-decoration: underline; }
    .toc ol { margin: 2px 0; padding-left: 18px; }
    .toc li { margin: 1px 0; }
    hr { border: none; border-top: 1px solid #ddd; margin: 15px 0; }
</style>
"""

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    df = pd.read_csv(DATA_PATH)

    totals = df[df["app_name"] == "--- TOTAL ---"].copy()
    mar = totals[totals["period"] == "4_Mar_MTD"].sort_values("avg_daily_spend", ascending=False)
    account_order = mar["account"].tolist()
    for a in df["account"].unique():
        if a not in account_order:
            account_order.append(a)

    accounts_str = ", ".join(ACCOUNT_SHORT.get(a, a) for a in account_order)

    html = [f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
    <title>P1 Assess: Spend Down SOI Down</title>{CSS}</head><body>"""]

    html.append("<h1>P1 Assess: SOI Down + Spend Down — Deep Dive</h1>")
    html.append(f'<p class="meta"><strong>Accounts:</strong> {accounts_str}<br>')
    html.append("<strong>Period:</strong> Benchmark (Dec 25-31, 2025) &rarr; Mar 1-13, 2026<br>")
    html.append("<strong>Criteria:</strong> Both avg daily spend AND SOI declined in March vs Benchmark<br>")
    html.append("<strong>Key question:</strong> Is this intentional budget reallocation, seasonal pullback, or early churn signal? For accounts with sharp Moloco install drops, check for campaign pauses, creative fatigue, or budget shifts to other channels.<br>")
    html.append("<strong>SOI:</strong> Account/app level = naive sum(Moloco) / sum(all installs). Multi-account = spend-weighted.</p>")

    # TOC
    html.append('<div class="toc"><strong>Contents</strong><ol>')
    for acct in account_order:
        short = ACCOUNT_SHORT.get(acct, acct)
        anchor = re.sub(r"[^\w\-]", "-", safe_filename(acct))
        html.append(f'<li><a href="#{anchor}">{short}</a></li>')
    html.append("</ol></div><hr>")

    for acct in account_order:
        short = ACCOUNT_SHORT.get(acct, acct)
        anchor = re.sub(r"[^\w\-]", "-", safe_filename(acct))
        acct_total = df[(df["account"] == acct) & (df["app_name"] == "--- TOTAL ---")].copy()
        acct_apps = df[(df["account"] == acct) & (df["app_name"] != "--- TOTAL ---")].copy()

        html.append(f'<h2 id="{anchor}">{short}</h2>')

        if not acct_total.empty:
            b64 = make_account_chart_b64(acct_total, short)
            html.append(f'<div class="chart"><img src="data:image/png;base64,{b64}" alt="{short}"></div>')
            html.append(html_account_table(acct_total))

        if not acct_apps.empty and not acct_total.empty:
            html.append("<h3>SOI Impact by App (Benchmark &rarr; Mar)</h3>")
            html.append(html_impact_table(acct_apps, acct_total))

        if not acct_apps.empty:
            html.append("<h3>App-Level SOI vs Spend</h3>")
            b64 = make_app_charts_b64(acct_apps, acct)
            if b64:
                html.append(f'<div class="chart"><img src="data:image/png;base64,{b64}" alt="{short} apps"></div>')
            html.append("<h3>App-Level Data</h3>")
            html.append(html_app_pivot(acct_apps))

        html.append("<hr>")

    html.append('<p class="meta"><em>Generated 2026-03-26. SOI = Moloco Installs / (Moloco + Non-Moloco Installs). DRR = Avg Daily Spend.</em></p>')
    html.append("</body></html>")

    path = os.path.join(OUTPUT_DIR, "p1_assess_deep_dive.html")
    with open(path, "w") as f:
        f.write("\n".join(html))
    print(f"Report: {path}")


if __name__ == "__main__":
    main()
