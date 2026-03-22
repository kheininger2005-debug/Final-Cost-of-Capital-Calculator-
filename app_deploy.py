
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd
import io, json, os
from datetime import datetime
from pdf_generator import generate_wacc_pdf

st.set_page_config(page_title="Cost of Capital Calculator", page_icon="📊", layout="wide")

st.markdown("""
<style>
.metric-card {background:linear-gradient(135deg,#1a1200,#0f0c00);border:1px solid #c8a84b55;border-radius:10px;padding:1.2rem 1.5rem;text-align:center;margin-bottom:1rem;}
.metric-label {font-size:0.65rem;letter-spacing:0.2em;text-transform:uppercase;color:#8b6914;margin-bottom:0.3rem;}
.metric-value {font-size:2rem;color:#f0d060;font-weight:700;}
.metric-value-sm {font-size:1.4rem;color:#c8a84b;font-weight:700;}
.summary-box {background:#0f0c00;border:1px solid #c8a84b33;border-left:4px solid #c8a84b;border-radius:8px;padding:1.5rem;margin:1rem 0;color:#d4b870;line-height:1.8;}
.summary-headline {font-size:1.1rem;color:#f0d060;margin-bottom:0.75rem;font-weight:700;}
.warning-box {background:#1a0f00;border:1px solid #ff9933;border-radius:6px;padding:0.75rem 1rem;color:#ffb366;font-size:0.8rem;margin:0.5rem 0;}
.bench-above {color:#ef5350;font-weight:700;}
.bench-below {color:#4caf50;font-weight:700;}
</style>
""", unsafe_allow_html=True)

INDUSTRY_BENCHMARKS = {
    "Technology": 9.5, "Healthcare": 8.2, "Financial Services": 9.8,
    "Consumer Discretionary": 8.7, "Consumer Staples": 6.9, "Energy": 10.4,
    "Utilities": 5.8, "Real Estate (REIT)": 7.2, "Industrials": 8.5,
    "Materials": 8.9, "Telecommunications": 7.6, "Retail": 8.1,
    "Automotive": 9.2, "Aerospace & Defense": 8.0, "Pharmaceuticals": 8.6,
    "Semiconductors": 10.1, "Media & Entertainment": 9.3, "Transportation": 8.4,
    "Food & Beverage": 7.0, "Insurance": 9.0
}

SAVE_FILE = "wacc_saved_analyses.json"

def load_saved():
    if os.path.exists(SAVE_FILE):
        with open(SAVE_FILE) as f:
            return json.load(f)
    return {}

def save_analysis(name, data):
    saved = load_saved()
    saved[name] = data
    with open(SAVE_FILE, "w") as f:
        json.dump(saved, f)

def calc_wacc(Re, Rd, Rp, T, We, Wd, Wp):
    after_tax_debt = Rd*(1-T)
    return We*Re + Wd*after_tax_debt + Wp*Rp, after_tax_debt

def calc_capm(Rf, beta, Rm):
    return Rf + beta*(Rm-Rf)

def get_risk_label(wacc):
    if wacc < 0.06: return "Conservative / Low-Risk", "#4caf50"
    if wacc < 0.10: return "Moderate", "#ffb300"
    if wacc < 0.14: return "Elevated", "#ff7043"
    return "High-Risk / Aggressive", "#ef5350"

def get_interpretation(company, wacc, Re, Rd, Rp, T, We, Wd, Wp, Rf, beta, Rm):
    after_tax = Rd*(1-T)
    wacc_pct = wacc*100
    risk_label, _ = get_risk_label(wacc)
    leverage = "heavily leveraged" if Wd>0.6 else ("moderately leveraged" if Wd>0.4 else "equity-heavy")
    name = company if company else "The firm"
    capm_re = calc_capm(Rf, beta, Rm) if Rm > 0 else None
    headline = f"{name}'s WACC of {wacc_pct:.2f}% reflects a {risk_label.lower()} capital structure."
    body = (f"At {wacc_pct:.2f}%, {name} must generate returns above this threshold on every dollar deployed. "
            f"The capital structure is {leverage}: equity at {We*100:.1f}% (cost {Re*100:.2f}%), "
            f"debt at {Wd*100:.1f}% (pre-tax cost {Rd*100:.2f}%).")
    if Wp > 0: body += f" Preferred stock: {Wp*100:.1f}% at {Rp*100:.2f}%."
    tax_shield = ""
    if Wd > 0:
        savings = (Rd-after_tax)*Wd*100
        tax_shield = (f"\n\nThe interest tax shield reduces effective cost of debt from {Rd*100:.2f}% "
                      f"to {after_tax*100:.2f}%, saving {savings:.2f}% in blended cost of capital.")
    capm_note = ""
    if capm_re is not None:
        diff = (Re-capm_re)*100
        capm_note = (f"\n\nCAPM (Rf={Rf*100:.2f}%, Beta={beta:.2f}, Rm={Rm*100:.2f}%) implies Re={capm_re*100:.2f}%. "
                     f"Input Re is {Re*100:.2f}%, difference of {diff:+.2f}% — ")
        capm_note += "aligned with market." if abs(diff)<0.5 else ("reflects company-specific risk." if diff>0 else "conservative estimate.")
    strategic = (f"\n\nAny project must clear {wacc_pct:.2f}% IRR to create value. "
                 + ("High WACC constrains investments — consider optimizing capital structure." if wacc>0.12
                    else "Current WACC supports value-accretive investment."))
    return headline, body+tax_shield+capm_note+strategic, capm_re

def run_monte_carlo(Re, Rd, T, We, Wd, Wp, Rp, beta, Rf, Rm, n=10000):
    np.random.seed(42)
    Re_sim  = np.random.normal(Re, Re*0.15, n)
    Rd_sim  = np.random.normal(Rd, Rd*0.10, n)
    bet_sim = np.random.normal(beta, beta*0.20, n) if Rm > 0 else np.full(n, beta)
    Re_sim  = np.where(Rm>0, Rf + bet_sim*(Rm-Rf), Re_sim)
    after_tax_sim = Rd_sim*(1-T)
    wacc_sim = We*Re_sim + Wd*after_tax_sim + Wp*Rp
    return wacc_sim*100

def find_optimal_structure(Re, Rd, T, Wp, Rp, base_Re=None):
    wd_range = np.linspace(0, 1-Wp, 200)
    waccs = []
    for wd in wd_range:
        we = 1 - wd - Wp
        rd_adj = Rd * (1 + wd*0.8)
        re_adj = Re * (1 + wd*0.5)
        w, _ = calc_wacc(re_adj, rd_adj, Rp, T, we, wd, Wp)
        waccs.append(w*100)
    opt_idx = np.argmin(waccs)
    return wd_range, np.array(waccs), wd_range[opt_idx], waccs[opt_idx]

def make_excel(r):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        gold   = wb.add_format({"bold":True,"font_color":"#c8a84b","bg_color":"#1a1200","border":1,"border_color":"#c8a84b44"})
        header = wb.add_format({"bold":True,"font_color":"#f0d060","bg_color":"#0f0c00","border":1,"border_color":"#c8a84b44"})
        normal = wb.add_format({"font_color":"#d4b870","bg_color":"#0a0800","border":1,"border_color":"#c8a84b22"})
        pct    = wb.add_format({"font_color":"#d4b870","bg_color":"#0a0800","num_format":"0.00%","border":1,"border_color":"#c8a84b22"})

        # Sheet 1: Summary
        ws = wb.add_worksheet("Summary")
        ws.set_tab_color("#c8a84b")
        ws.set_column("A:A", 30); ws.set_column("B:B", 20); ws.set_column("C:C", 35)
        ws.write("A1", "COST OF CAPITAL REPORT", header); ws.write("B1", r["company"] or "Analysis", header)
        ws.write("A2", "Generated", normal); ws.write("B2", r["timestamp"], normal)
        ws.write("A4", "PARAMETER", header); ws.write("B4", "VALUE", header); ws.write("C4", "NOTES", header)
        rows = [
            ("Cost of Equity (Re)", r["Re"], "Required return by equity holders"),
            ("Cost of Debt (Rd)", r["Rd"], "Pre-tax rate on debt"),
            ("After-Tax Cost of Debt", r["after_tax_debt"], f"Rd x (1-T)"),
            ("Cost of Preferred (Rp)", r["Rp"], "Preferred dividend yield"),
            ("Corporate Tax Rate (T)", r["T"], "Marginal tax rate"),
            ("Weight of Equity (We)", r["We"], "Equity % of total capital"),
            ("Weight of Debt (Wd)", r["Wd"], "Debt % of total capital"),
            ("Weight of Preferred (Wp)", r["Wp"], "Preferred % of total capital"),
            ("WACC", r["wacc"], "Weighted Average Cost of Capital"),
        ]
        if r.get("capm_re"):
            rows += [("Risk-Free Rate (Rf)", r["Rf"], ""), ("Beta", r["beta"], "Systematic risk"),
                     ("Market Return (Rm)", r["Rm"], ""), ("CAPM-Implied Re", r["capm_re"], "Rf+Beta*(Rm-Rf)")]
        for i, (param, val, note) in enumerate(rows, start=4):
            ws.write(i, 0, param, gold if param=="WACC" else normal)
            ws.write(i, 1, val, pct)
            ws.write(i, 2, note, normal)

        # Sheet 2: Monte Carlo
        ws2 = wb.add_worksheet("Monte Carlo")
        ws2.set_tab_color("#5b9bd5")
        wacc_sim = run_monte_carlo(r["Re"],r["Rd"],r["T"],r["We"],r["Wd"],r["Wp"],r["Rp"],r["beta"],r["Rf"],r["Rm"])
        ws2.write("A1","Monte Carlo Simulation Results",header)
        ws2.write("A3","Metric",header); ws2.write("B3","Value",header)
        stats = [("Mean WACC",f"{np.mean(wacc_sim):.2f}%"),("Median WACC",f"{np.median(wacc_sim):.2f}%"),
                 ("Std Dev",f"{np.std(wacc_sim):.2f}%"),("5th Percentile",f"{np.percentile(wacc_sim,5):.2f}%"),
                 ("95th Percentile",f"{np.percentile(wacc_sim,95):.2f}%"),("Min",f"{np.min(wacc_sim):.2f}%"),("Max",f"{np.max(wacc_sim):.2f}%")]
        for i,(k,v) in enumerate(stats,start=3):
            ws2.write(i,0,k,normal); ws2.write(i,1,v,normal)

        # Sheet 3: DCF
        ws3 = wb.add_worksheet("DCF Valuation")
        ws3.set_tab_color("#9b59b6")
        ws3.write("A1","DCF Valuation Summary",header)
        ws3.write("A3","Inputs",header)

    buf.seek(0)
    return buf.read()

# ── Sidebar ──
with st.sidebar:
    st.markdown("## 📊 Cost of Capital")

    # Save/Load
    saved = load_saved()
    if saved:
        st.markdown("**💾 Load Saved Analysis**")
        load_choice = st.selectbox("Select", ["-- New Analysis --"] + list(saved.keys()))
        if load_choice != "-- New Analysis --" and st.button("Load"):
            st.session_state.wacc_result = saved[load_choice]
            st.success(f"Loaded: {load_choice}")

    st.markdown("---")
    company_name = st.text_input("Company Name", placeholder="e.g. Apple Inc.")
    industry = st.selectbox("Industry (for benchmark)", ["-- Select --"] + list(INDUSTRY_BENCHMARKS.keys()))

    st.markdown("**Capital Costs**")
    Re = st.number_input("Cost of Equity — Re (%)", 0.0, 100.0, 10.0, 0.1) / 100
    Rd = st.number_input("Cost of Debt — Rd (%)", 0.0, 100.0, 5.0, 0.1) / 100
    Rp = st.number_input("Cost of Preferred — Rp (%)", 0.0, 100.0, 0.0, 0.1) / 100
    T  = st.number_input("Tax Rate (%)", 0.0, 100.0, 21.0, 0.5) / 100

    st.markdown("**Capital Weights** *(must sum to 100%)*")
    We = st.number_input("Weight of Equity (%)", 0.0, 100.0, 60.0, 1.0) / 100
    Wd = st.number_input("Weight of Debt (%)", 0.0, 100.0, 40.0, 1.0) / 100
    Wp = st.number_input("Weight of Preferred (%)", 0.0, 100.0, 0.0, 1.0) / 100
    weight_sum = (We+Wd+Wp)*100
    if abs(weight_sum-100) > 0.5:
        st.markdown(f'<div class="warning-box">⚠️ Weights sum to {weight_sum:.1f}%</div>', unsafe_allow_html=True)

    st.markdown("**CAPM Inputs** *(optional)*")
    Rf   = st.number_input("Risk-Free Rate — Rf (%)", 0.0, 20.0, 0.0, 0.1) / 100
    beta = st.number_input("Equity Beta (β)", 0.0, 5.0, 1.0, 0.05)
    Rm   = st.number_input("Market Return — Rm (%)", 0.0, 30.0, 0.0, 0.1) / 100

    compute_btn = st.button("▶ COMPUTE WACC", use_container_width=True)

    # Save button
    if "wacc_result" in st.session_state:
        st.markdown("---")
        save_name = st.text_input("Save as...", value=company_name or "Analysis")
        if st.button("💾 Save Analysis"):
            save_analysis(save_name, st.session_state.wacc_result)
            st.success(f"Saved: {save_name}")

# ── Main ──
st.markdown("<h1 style='text-align:center;color:#f0d060;'> Cost of Capital </h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:#8b6914;'>WACC · CAPM · Monte Carlo · Benchmarks · DCF · Company Comparison</p>", unsafe_allow_html=True)
st.markdown("---")

if compute_btn or "wacc_result" in st.session_state:
    if compute_btn:
        if abs((We+Wd+Wp)-1.0) > 0.005:
            st.error(f"⚠️ Weights must sum to 100%. Currently: {(We+Wd+Wp)*100:.1f}%")
            st.stop()
        wacc, after_tax_debt = calc_wacc(Re, Rd, Rp, T, We, Wd, Wp)
        capm_re = calc_capm(Rf, beta, Rm) if Rm > 0 else None
        headline, summary, _ = get_interpretation(company_name, wacc, Re, Rd, Rp, T, We, Wd, Wp, Rf, beta, Rm)
        st.session_state.wacc_result = {
            "wacc":wacc,"Re":Re,"Rd":Rd,"Rp":Rp,"T":T,"We":We,"Wd":Wd,"Wp":Wp,
            "Rf":Rf,"beta":beta,"Rm":Rm,"after_tax_debt":after_tax_debt,
            "capm_re":capm_re,"headline":headline,"summary":summary,
            "company":company_name,"industry":industry,"timestamp":datetime.now().strftime("%B %d, %Y")
        }

    r = st.session_state.wacc_result
    wacc = r["wacc"]
    risk_label, risk_color = get_risk_label(wacc)
    bench = INDUSTRY_BENCHMARKS.get(r.get("industry",""), None)

    # ── Metric Cards ──
    cols = st.columns(5)
    cards = [
        ("WACC", f"{wacc*100:.2f}%", "metric-value"),
        ("Cost of Equity", f"{r['Re']*100:.2f}%", "metric-value-sm"),
        ("After-Tax Debt", f"{r['after_tax_debt']*100:.2f}%", "metric-value-sm"),
        ("Risk Profile", risk_label, "metric-value-sm"),
        ("vs Industry" if bench else "CAPM Re",
         (f"{'▲' if wacc*100>bench else '▼'} {abs(wacc*100-bench):.2f}%" if bench
          else (f"{r['capm_re']*100:.2f}%" if r.get('capm_re') else "N/A")),
         "metric-value-sm"),
    ]
    for col, (label, val, cls) in zip(cols, cards):
        with col:
            color = risk_color if label=="Risk Profile" else ("#ef5350" if (bench and wacc*100>bench and label=="vs Industry") else ("#4caf50" if (bench and label=="vs Industry") else "#c8a84b"))
            st.markdown(f'<div class="metric-card"><div class="metric-label">{label}</div><div class="{cls}" style="color:{color};">{val}</div></div>', unsafe_allow_html=True)

    # ── Tabs ──
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📈 Charts", "🏭 Benchmark", "🎲 Monte Carlo",
        "⚖️ Optimal Structure", "🏢 Compare Companies",
        "💹 DCF Valuation", "📥 Export"
    ])

    # TAB 1: Charts
    with tab1:
        fig, axes = plt.subplots(2, 2, figsize=(12, 8))
        fig.patch.set_facecolor("#0a0800")
        for ax in axes.flat:
            ax.set_facecolor("#0f0c00")
            for sp in ["top","right"]: ax.spines[sp].set_visible(False)
            ax.spines["bottom"].set_color("#c8a84b33"); ax.spines["left"].set_color("#c8a84b33")
            ax.tick_params(colors="#a08040", labelsize=8)
        comps,labs,cols2=[],[],[]
        if r["We"]>0: comps.append(r["We"]*r["Re"]*100); labs.append("Equity"); cols2.append("#c8a84b")
        if r["Wd"]>0: comps.append(r["Wd"]*r["after_tax_debt"]*100); labs.append("Debt (after-tax)"); cols2.append("#5b9bd5")
        if r["Wp"]>0: comps.append(r["Wp"]*r["Rp"]*100); labs.append("Preferred"); cols2.append("#9b59b6")
        bars = axes[0,0].barh(labs, comps, color=cols2, height=0.45, edgecolor="none")
        for bar,val in zip(bars,comps):
            axes[0,0].text(bar.get_width()+0.01,bar.get_y()+bar.get_height()/2,f"{val:.3f}%",va="center",color="#c8a84b",fontsize=9,fontweight="bold")
        axes[0,0].set_title("WACC Components", color="#f0d060", fontsize=11, fontweight="bold")
        axes[0,0].set_xlabel("Contribution (%)", color="#8b6914", fontsize=9)
        axes[0,1].set_facecolor("#0a0800")
        sizes2=[r["We"]*100,r["Wd"]*100]; labs2=[f"Equity\n{r['We']*100:.1f}%",f"Debt\n{r['Wd']*100:.1f}%"]; cols3=["#c8a84b","#5b9bd5"]
        if r["Wp"]>0: sizes2.append(r["Wp"]*100); labs2.append(f"Preferred\n{r['Wp']*100:.1f}%"); cols3.append("#9b59b6")
        axes[0,1].pie(sizes2,labels=labs2,colors=cols3,startangle=90,wedgeprops=dict(edgecolor="#0a0800",linewidth=2),textprops=dict(color="#c8a84b",fontsize=9))
        axes[0,1].set_title("Capital Structure", color="#f0d060", fontsize=11, fontweight="bold")
        re_range=np.linspace(max(0.01,r["Re"]-0.06),r["Re"]+0.06,50)
        waccs_re=[r["We"]*rv+r["Wd"]*r["Rd"]*(1-r["T"])+r["Wp"]*r["Rp"] for rv in re_range]
        axes[1,0].plot(re_range*100,[w*100 for w in waccs_re],color="#c8a84b",lw=2)
        axes[1,0].axvline(r["Re"]*100,color="#c8a84b",ls="--",alpha=0.4)
        axes[1,0].fill_between(re_range*100,[w*100 for w in waccs_re],alpha=0.1,color="#c8a84b")
        axes[1,0].set_title("WACC vs Cost of Equity",color="#f0d060",fontsize=11,fontweight="bold")
        axes[1,0].set_xlabel("Cost of Equity (%)",color="#8b6914",fontsize=9); axes[1,0].set_ylabel("WACC (%)",color="#8b6914",fontsize=9)
        wd_r=np.linspace(0,1-r["Wp"],50); we_r=1-wd_r-r["Wp"]
        waccs2=[max(0,we)*r["Re"]+wd*r["Rd"]*(1-r["T"])+r["Wp"]*r["Rp"] for we,wd in zip(we_r,wd_r)]
        axes[1,1].plot(wd_r*100,[w*100 for w in waccs2],color="#5b9bd5",lw=2)
        axes[1,1].axvline(r["Wd"]*100,color="#5b9bd5",ls="--",alpha=0.4)
        axes[1,1].fill_between(wd_r*100,[w*100 for w in waccs2],alpha=0.1,color="#5b9bd5")
        axes[1,1].set_title("WACC vs Leverage",color="#f0d060",fontsize=11,fontweight="bold")
        axes[1,1].set_xlabel("Weight of Debt (%)",color="#8b6914",fontsize=9); axes[1,1].set_ylabel("WACC (%)",color="#8b6914",fontsize=9)
        plt.tight_layout(); st.pyplot(fig); plt.close(fig)

    # TAB 2: Benchmark
    with tab2:
        st.markdown("### 🏭 Industry Benchmark Comparison")
        fig2, ax2 = plt.subplots(figsize=(12,5))
        fig2.patch.set_facecolor("#0a0800"); ax2.set_facecolor("#0f0c00")
        industries = list(INDUSTRY_BENCHMARKS.keys())
        values = list(INDUSTRY_BENCHMARKS.values())
        bar_colors = ["#ef5350" if v < wacc*100 else "#4caf50" for v in values]
        bars2 = ax2.barh(industries, values, color=bar_colors, height=0.6, edgecolor="none", alpha=0.7)
        ax2.axvline(wacc*100, color="#f0d060", linewidth=2.5, linestyle="--", label=f"{r['company'] or 'Your WACC'}: {wacc*100:.2f}%")
        ax2.set_xlabel("WACC (%)", color="#8b6914", fontsize=10)
        ax2.set_title("Your WACC vs Industry Averages", color="#f0d060", fontsize=13, fontweight="bold")
        ax2.tick_params(colors="#a08040", labelsize=9)
        for sp in ["top","right"]: ax2.spines[sp].set_visible(False)
        ax2.spines["bottom"].set_color("#c8a84b33"); ax2.spines["left"].set_color("#c8a84b33")
        ax2.legend(fontsize=10, facecolor="#1a1200", edgecolor="#c8a84b44", labelcolor="#f0d060")
        plt.tight_layout(); st.pyplot(fig2); plt.close(fig2)
        above = [(k,v) for k,v in INDUSTRY_BENCHMARKS.items() if wacc*100 > v]
        below = [(k,v) for k,v in INDUSTRY_BENCHMARKS.items() if wacc*100 <= v]
        c1,c2 = st.columns(2)
        with c1:
            st.markdown(f"** Higher than {len(above)} industries:**")
            for k,v in above: st.markdown(f"- {k}: {v:.1f}% (yours is +{wacc*100-v:.2f}%)")
        with c2:
            st.markdown(f"**Lower than {len(below)} industries:**")
            for k,v in below: st.markdown(f"- {k}: {v:.1f}% (yours is -{v-wacc*100:.2f}%)")

    # TAB 3: Monte Carlo
    with tab3:
        st.markdown("### 🎲 Monte Carlo Simulation (10,000 runs)")
        st.markdown("*Varies Re ±15%, Rd ±10%, and Beta ±20% to show WACC uncertainty range.*")
        wacc_sim = run_monte_carlo(r["Re"],r["Rd"],r["T"],r["We"],r["Wd"],r["Wp"],r["Rp"],r["beta"],r["Rf"],r["Rm"])
        fig3, ax3 = plt.subplots(figsize=(10,4))
        fig3.patch.set_facecolor("#0a0800"); ax3.set_facecolor("#0f0c00")
        ax3.hist(wacc_sim, bins=80, color="#c8a84b", edgecolor="none", alpha=0.8)
        ax3.axvline(np.mean(wacc_sim),color="#f0d060",lw=2,linestyle="--",label=f"Mean: {np.mean(wacc_sim):.2f}%")
        ax3.axvline(np.percentile(wacc_sim,5),color="#5b9bd5",lw=1.5,linestyle=":",label=f"5th pct: {np.percentile(wacc_sim,5):.2f}%")
        ax3.axvline(np.percentile(wacc_sim,95),color="#ef5350",lw=1.5,linestyle=":",label=f"95th pct: {np.percentile(wacc_sim,95):.2f}%")
        ax3.axvline(wacc*100,color="#ffffff",lw=2,label=f"Your WACC: {wacc*100:.2f}%")
        ax3.set_xlabel("WACC (%)",color="#8b6914",fontsize=10); ax3.set_ylabel("Frequency",color="#8b6914",fontsize=10)
        ax3.set_title("WACC Distribution under Uncertainty",color="#f0d060",fontsize=13,fontweight="bold")
        ax3.tick_params(colors="#a08040"); ax3.legend(facecolor="#1a1200",edgecolor="#c8a84b44",labelcolor="#f0d060",fontsize=9)
        for sp in ["top","right"]: ax3.spines[sp].set_visible(False)
        ax3.spines["bottom"].set_color("#c8a84b33"); ax3.spines["left"].set_color("#c8a84b33")
        plt.tight_layout(); st.pyplot(fig3); plt.close(fig3)
        mc1,mc2,mc3,mc4 = st.columns(4)
        for col,label,val in zip([mc1,mc2,mc3,mc4],
            ["Mean WACC","Std Dev","5th Percentile","95th Percentile"],
            [f"{np.mean(wacc_sim):.2f}%",f"{np.std(wacc_sim):.2f}%",f"{np.percentile(wacc_sim,5):.2f}%",f"{np.percentile(wacc_sim,95):.2f}%"]):
            with col:
                st.markdown(f'<div class="metric-card"><div class="metric-label">{label}</div><div class="metric-value-sm">{val}</div></div>', unsafe_allow_html=True)

    # TAB 4: Optimal Structure
    with tab4:
        st.markdown("### ⚖️ Optimal Capital Structure")
        st.markdown("*Shows how WACC changes as leverage increases, accounting for rising risk at higher debt levels.*")
        wd_range, wacc_curve, opt_wd, opt_wacc = find_optimal_structure(r["Re"],r["Rd"],r["T"],r["Wp"],r["Rp"])
        fig4, ax4 = plt.subplots(figsize=(10,4))
        fig4.patch.set_facecolor("#0a0800"); ax4.set_facecolor("#0f0c00")
        ax4.plot(wd_range*100, wacc_curve, color="#c8a84b", lw=2.5)
        ax4.axvline(opt_wd*100,color="#4caf50",lw=2,linestyle="--",label=f"Optimal Wd: {opt_wd*100:.1f}% → WACC: {opt_wacc:.2f}%")
        ax4.axvline(r["Wd"]*100,color="#f0d060",lw=1.5,linestyle=":",label=f"Current Wd: {r['Wd']*100:.1f}% → WACC: {wacc*100:.2f}%")
        ax4.fill_between(wd_range*100, wacc_curve, alpha=0.1, color="#c8a84b")
        ax4.scatter([opt_wd*100],[opt_wacc],color="#4caf50",s=100,zorder=5)
        ax4.set_xlabel("Weight of Debt (%)",color="#8b6914",fontsize=10); ax4.set_ylabel("WACC (%)",color="#8b6914",fontsize=10)
        ax4.set_title("WACC vs Capital Structure (with Risk Adjustment)",color="#f0d060",fontsize=13,fontweight="bold")
        ax4.tick_params(colors="#a08040"); ax4.legend(facecolor="#1a1200",edgecolor="#c8a84b44",labelcolor="#f0d060",fontsize=9)
        for sp in ["top","right"]: ax4.spines[sp].set_visible(False)
        ax4.spines["bottom"].set_color("#c8a84b33"); ax4.spines["left"].set_color("#c8a84b33")
        plt.tight_layout(); st.pyplot(fig4); plt.close(fig4)
        gap = wacc*100 - opt_wacc
        if gap > 0.1:
            st.success(f"Moving to optimal structure ({opt_wd*100:.1f}% debt) could reduce WACC by **{gap:.2f}%**, increasing firm value.")
        else:
            st.success(f" Your current structure is near-optimal! You are only {abs(gap):.2f}% from the minimum WACC.")

    # TAB 5: Company Comparison
    with tab5:
        st.markdown("###  Compare Multiple Companies")
        st.markdown("Enter data for up to 2 additional companies to compare side-by-side.")
        companies = [{"name": r["company"] or "Company 1", "wacc": wacc*100, "Re": r["Re"]*100, "Rd": r["Rd"]*100, "We": r["We"]*100, "Wd": r["Wd"]*100}]
        for i in range(2):
            with st.expander(f"➕ Company {i+2}", expanded=(i==0)):
                cn = st.text_input(f"Name", key=f"cn{i}", value=f"Company {i+2}")
                c1,c2,c3,c4,c5 = st.columns(5)
                cRe = c1.number_input("Re (%)",0.0,50.0,10.0,0.1,key=f"cRe{i}")
                cRd = c2.number_input("Rd (%)",0.0,30.0,5.0,0.1,key=f"cRd{i}")
                cT  = c3.number_input("T (%)",0.0,50.0,21.0,0.5,key=f"cT{i}")
                cWe = c4.number_input("We (%)",0.0,100.0,60.0,1.0,key=f"cWe{i}")
                cWd = c5.number_input("Wd (%)",0.0,100.0,40.0,1.0,key=f"cWd{i}")
                if abs(cWe+cWd-100)<1:
                    cwacc = cWe/100*cRe/100 + cWd/100*(cRd/100)*(1-cT/100)
                    companies.append({"name":cn,"wacc":cwacc*100,"Re":cRe,"Rd":cRd,"We":cWe,"Wd":cWd})
        if len(companies) > 1:
            fig5, axes5 = plt.subplots(1,2,figsize=(12,4))
            fig5.patch.set_facecolor("#0a0800")
            for ax in axes5:
                ax.set_facecolor("#0f0c00")
                for sp in ["top","right"]: ax.spines[sp].set_visible(False)
                ax.spines["bottom"].set_color("#c8a84b33"); ax.spines["left"].set_color("#c8a84b33")
                ax.tick_params(colors="#a08040",labelsize=9)
            names=[c["name"] for c in companies]; waccs_c=[c["wacc"] for c in companies]
            bar_c=["#c8a84b","#5b9bd5","#9b59b6"][:len(companies)]
            axes5[0].bar(names,waccs_c,color=bar_c,edgecolor="none",width=0.5)
            for i,(n,v) in enumerate(zip(names,waccs_c)):
                axes5[0].text(i,v+0.05,f"{v:.2f}%",ha="center",color="#f0d060",fontsize=10,fontweight="bold")
            axes5[0].set_ylabel("WACC (%)",color="#8b6914",fontsize=9); axes5[0].set_title("WACC Comparison",color="#f0d060",fontsize=11,fontweight="bold")
            x=np.arange(len(companies)); width=0.25
            axes5[1].bar(x-width,[c["Re"] for c in companies],width,label="Re",color="#c8a84b",edgecolor="none")
            axes5[1].bar(x,[c["Rd"] for c in companies],width,label="Rd",color="#5b9bd5",edgecolor="none")
            axes5[1].bar(x+width,[c["We"] for c in companies],width,label="We%",color="#9b59b6",edgecolor="none",alpha=0.7)
            axes5[1].set_xticks(x); axes5[1].set_xticklabels(names,color="#a08040")
            axes5[1].set_title("Cost & Weight Comparison",color="#f0d060",fontsize=11,fontweight="bold")
            axes5[1].legend(facecolor="#1a1200",edgecolor="#c8a84b44",labelcolor="#f0d060",fontsize=8)
            plt.tight_layout(); st.pyplot(fig5); plt.close(fig5)
            df_comp = pd.DataFrame(companies).round(2)
            st.dataframe(df_comp, use_container_width=True, hide_index=True)

    # TAB 6: DCF Valuation
    with tab6:
        st.markdown("### 💹 DCF Valuation (WACC as Discount Rate)")
        st.markdown("*Use your computed WACC to estimate firm value via discounted cash flow.*")
        d1,d2,d3 = st.columns(3)
        fcf     = d1.number_input("Free Cash Flow — Year 1 ($M)", 0.0, 100000.0, 100.0, 1.0)
        g_near  = d2.number_input("Near-term Growth Rate (%)", 0.0, 50.0, 10.0, 0.5) / 100
        g_term  = d3.number_input("Terminal Growth Rate (%)", 0.0, 10.0, 2.5, 0.1) / 100
        years   = st.slider("Projection Period (years)", 3, 15, 5)

        if wacc > g_term:
            fcfs = [fcf*(1+g_near)**i for i in range(1, years+1)]
            disc = [f/((1+wacc)**i) for i,f in enumerate(fcfs,1)]
            tv   = fcfs[-1]*(1+g_term)/(wacc-g_term)
            tv_pv= tv/((1+wacc)**years)
            firm_val = sum(disc)+tv_pv
            fig6, ax6 = plt.subplots(figsize=(10,4))
            fig6.patch.set_facecolor("#0a0800"); ax6.set_facecolor("#0f0c00")
            yr_labels=[f"Y{i}" for i in range(1,years+1)]
            ax6.bar(yr_labels, disc, color="#c8a84b", edgecolor="none", label="PV of FCF", width=0.5)
            ax6.bar(["Terminal"], [tv_pv], color="#5b9bd5", edgecolor="none", label=f"PV Terminal Value", width=0.5)
            for i,(lbl,val) in enumerate(zip(yr_labels+["Terminal"],disc+[tv_pv])):
                ax6.text(i,val+firm_val*0.005,f"${val:.0f}M",ha="center",color="#f0d060",fontsize=8,fontweight="bold")
            ax6.set_ylabel("Present Value ($M)",color="#8b6914",fontsize=9)
            ax6.set_title(f"DCF Components — Firm Value: ${firm_val:.1f}M",color="#f0d060",fontsize=12,fontweight="bold")
            ax6.tick_params(colors="#a08040"); ax6.legend(facecolor="#1a1200",edgecolor="#c8a84b44",labelcolor="#f0d060",fontsize=9)
            for sp in ["top","right"]: ax6.spines[sp].set_visible(False)
            ax6.spines["bottom"].set_color("#c8a84b33"); ax6.spines["left"].set_color("#c8a84b33")
            plt.tight_layout(); st.pyplot(fig6); plt.close(fig6)
            v1,v2,v3 = st.columns(3)
            v1.markdown(f'<div class="metric-card"><div class="metric-label">PV of FCFs</div><div class="metric-value-sm">${sum(disc):.1f}M</div></div>', unsafe_allow_html=True)
            v2.markdown(f'<div class="metric-card"><div class="metric-label">PV Terminal Value</div><div class="metric-value-sm">${tv_pv:.1f}M</div></div>', unsafe_allow_html=True)
            v3.markdown(f'<div class="metric-card"><div class="metric-label">Total Firm Value</div><div class="metric-value">${firm_val:.1f}M</div></div>', unsafe_allow_html=True)
        else:
            st.error("⚠️ Terminal growth rate must be less than WACC for DCF to work.")

    # TAB 7: Export
    with tab7:
        st.markdown("### 📥 Export Your Analysis")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.markdown("**📄 AI-Powered PDF Report**")
            api_key_input = st.text_input("Anthropic API Key", type="password", placeholder="sk-ant-...", key="pdf_api_key")
            st.caption("Leave blank for template summary. Get a key at console.anthropic.com")
            btn1, btn2 = st.columns(2)
            with btn1:
                st.markdown("🤖 **AI PDF**")
                if api_key_input:
                    with st.spinner("Generating AI summary..."):
                        pdf_ai = generate_wacc_pdf(r, api_key=api_key_input)
                    st.download_button("⬇ Download AI PDF", data=pdf_ai,
                        file_name=f"WACC_AI_{(r['company'] or 'Report').replace(' ','_')}.pdf",
                        mime="application/pdf", key="dl_ai_pdf")
                else:
                    st.warning("Enter API key to enable")
            with btn2:
                st.markdown("📄 **Standard PDF**")
                with st.spinner("Building PDF..."):
                    pdf_std = generate_wacc_pdf(r, api_key=None)
                st.download_button("⬇ Download Standard PDF", data=pdf_std,
                    file_name=f"WACC_{(r['company'] or 'Report').replace(' ','_')}.pdf",
                    mime="application/pdf", key="dl_std_pdf")
        with col_b:
            st.markdown("**📊 Excel Workbook**")
            st.markdown("*Summary, Monte Carlo stats, DCF data*")
            with st.spinner("Building Excel..."): xl_bytes = make_excel(r)
            st.download_button("⬇ Download Excel", data=xl_bytes,
                file_name=f"WACC_{(r['company'] or 'Report').replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.markdown("**🔗 Share as Link**")
            st.markdown("*Deploy to Streamlit Cloud for a public URL*")
            st.info("1. Push files to GitHub\n2. Go to share.streamlit.io\n3. Connect repo → app.py\n4. Get a public link!")

else:
    st.markdown("<div style='text-align:center;padding:4rem 0;color:#5a4520;'><div style='font-size:3rem;'>◈</div><p>Enter inputs in the sidebar and click COMPUTE WACC</p></div>", unsafe_allow_html=True)
