
import io
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus.flowables import Flowable

GOLD       = colors.HexColor("#c8a84b")
GOLD_DARK  = colors.HexColor("#8b6914")
GOLD_LIGHT = colors.HexColor("#f0d060")
BG_DARK    = colors.HexColor("#0a0800")
BG_MID     = colors.HexColor("#0f0c00")
BG_CARD    = colors.HexColor("#1a1200")
TEXT_DIM   = colors.HexColor("#a08040")
TEXT_BODY  = colors.HexColor("#d4b870")
BLUE       = colors.HexColor("#5b9bd5")
W, H = A4

INDUSTRY_BENCHMARKS = {
    "Technology": 9.5, "Healthcare": 8.2, "Financial Services": 9.8,
    "Consumer Discretionary": 8.7, "Consumer Staples": 6.9, "Energy": 10.4,
    "Utilities": 5.8, "Real Estate (REIT)": 7.2, "Industrials": 8.5,
    "Materials": 8.9, "Telecommunications": 7.6, "Retail": 8.1,
    "Automotive": 9.2, "Aerospace & Defense": 8.0, "Pharmaceuticals": 8.6,
    "Semiconductors": 10.1, "Media & Entertainment": 9.3,
    "Transportation": 8.4, "Food & Beverage": 7.0, "Insurance": 9.0
}

class GoldRule(Flowable):
    def __init__(self, width, thickness=1.5):
        super().__init__()
        self.width = width
        self.height = thickness + 2
        self.thickness = thickness
    def draw(self):
        self.canv.setStrokeColor(GOLD)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self.width, 0)

def build_styles():
    styles = getSampleStyleSheet()
    def add(name, **kw):
        styles.add(ParagraphStyle(name=name, **kw))
    add("DocTitle",    fontName="Helvetica-Bold", fontSize=30, textColor=GOLD_LIGHT, alignment=TA_CENTER, spaceAfter=6, leading=36)
    add("DocSubtitle", fontName="Helvetica",      fontSize=13, textColor=GOLD,       alignment=TA_CENTER, spaceAfter=4, leading=16)
    add("DocCompany",  fontName="Helvetica-Bold", fontSize=18, textColor=colors.white, alignment=TA_CENTER, spaceAfter=8, leading=22)
    add("WACCHero",    fontName="Helvetica-Bold", fontSize=42, textColor=GOLD_LIGHT, alignment=TA_CENTER, leading=50)
    add("WACCLabel",   fontName="Helvetica",      fontSize=9,  textColor=TEXT_DIM,   alignment=TA_CENTER, leading=12)
    add("SectionTag",  fontName="Helvetica-Bold", fontSize=7,  textColor=GOLD_DARK,  spaceAfter=4, leading=10)
    add("Headline",    fontName="Helvetica-Bold", fontSize=13, textColor=GOLD_LIGHT, spaceAfter=8, leading=18)
    add("Body",        fontName="Helvetica",      fontSize=9.5,textColor=TEXT_BODY,  spaceAfter=6, leading=16)
    add("BodyBold",    fontName="Helvetica-Bold", fontSize=9.5,textColor=GOLD,       spaceAfter=4, leading=16)
    add("DateText",    fontName="Helvetica",      fontSize=8,  textColor=TEXT_DIM,   alignment=TA_CENTER)
    return styles

def fig_to_image(fig, width_pts, height_pts):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="#0a0800", edgecolor="none")
    buf.seek(0)
    plt.close(fig)
    return Image(buf, width=width_pts, height=height_pts)

def make_bar_chart(We, Re, Wd, after_tax, Wp, Rp, wacc):
    fig, ax = plt.subplots(figsize=(7, 3))
    fig.patch.set_facecolor("#0a0800"); ax.set_facecolor("#0f0c00")
    comps, labs, cols = [], [], []
    if We > 0: comps.append(We*Re*100); labs.append(f"Equity ({We*100:.0f}% x {Re*100:.2f}%)"); cols.append("#c8a84b")
    if Wd > 0: comps.append(Wd*after_tax*100); labs.append(f"Debt after-tax ({Wd*100:.0f}% x {after_tax*100:.2f}%)"); cols.append("#5b9bd5")
    if Wp > 0: comps.append(Wp*Rp*100); labs.append(f"Preferred ({Wp*100:.0f}% x {Rp*100:.2f}%)"); cols.append("#9b59b6")
    bars = ax.barh(labs, comps, color=cols, height=0.45, edgecolor="none")
    for bar, val in zip(bars, comps):
        ax.text(bar.get_width()+0.015, bar.get_y()+bar.get_height()/2, f"{val:.3f}%", va="center", color="#c8a84b", fontsize=8, fontweight="bold")
    ax.axvline(wacc*100, color="#c8a84b", linewidth=1.2, linestyle="--", alpha=0.5)
    ax.set_xlabel("Contribution to WACC (%)", color="#8b6914", fontsize=8)
    ax.tick_params(colors="#a08040", labelsize=7.5)
    for sp in ["top","right"]: ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color("#c8a84b33"); ax.spines["left"].set_color("#c8a84b33")
    ax.set_title("WACC Component Contributions", color="#c8a84b", fontsize=10, fontweight="bold", pad=8)
    plt.tight_layout(); return fig

def make_pie_chart(We, Wd, Wp):
    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    fig.patch.set_facecolor("#0a0800"); ax.set_facecolor("#0a0800")
    sizes, labs, cols = [], [], []
    if We > 0: sizes.append(We*100); labs.append(f"Equity {We*100:.1f}%"); cols.append("#c8a84b")
    if Wd > 0: sizes.append(Wd*100); labs.append(f"Debt {Wd*100:.1f}%"); cols.append("#5b9bd5")
    if Wp > 0: sizes.append(Wp*100); labs.append(f"Preferred {Wp*100:.1f}%"); cols.append("#9b59b6")
    ax.pie(sizes, labels=labs, colors=cols, startangle=90,
           wedgeprops=dict(edgecolor="#0a0800", linewidth=2),
           textprops=dict(color="#c8a84b", fontsize=8))
    ax.set_title("Capital Structure", color="#c8a84b", fontsize=10, fontweight="bold", pad=8)
    plt.tight_layout(); return fig

def make_sensitivity(Re, Rd, T, We, Wd, Wp, Rp):
    fig, axes = plt.subplots(1, 2, figsize=(9, 3.2))
    fig.patch.set_facecolor("#0a0800")
    for ax in axes:
        ax.set_facecolor("#0f0c00")
        for sp in ["top","right"]: ax.spines[sp].set_visible(False)
        ax.spines["bottom"].set_color("#c8a84b33"); ax.spines["left"].set_color("#c8a84b33")
        ax.tick_params(colors="#a08040", labelsize=7.5)
    re_range = np.linspace(max(0.01, Re-0.06), Re+0.06, 50)
    waccs = [We*r + Wd*Rd*(1-T) + Wp*Rp for r in re_range]
    axes[0].plot(re_range*100, [w*100 for w in waccs], color="#c8a84b", lw=2)
    axes[0].axvline(Re*100, color="#c8a84b", ls="--", alpha=0.4)
    axes[0].fill_between(re_range*100, [w*100 for w in waccs], alpha=0.1, color="#c8a84b")
    axes[0].set_xlabel("Cost of Equity (%)", color="#8b6914", fontsize=8)
    axes[0].set_ylabel("WACC (%)", color="#8b6914", fontsize=8)
    axes[0].set_title("WACC vs Cost of Equity", color="#c8a84b", fontsize=9, fontweight="bold")
    wd_r = np.linspace(0, 1-Wp, 50); we_r = 1 - wd_r - Wp
    waccs2 = [max(0,we)*Re + wd*Rd*(1-T) + Wp*Rp for we,wd in zip(we_r,wd_r)]
    axes[1].plot(wd_r*100, [w*100 for w in waccs2], color="#5b9bd5", lw=2)
    axes[1].axvline(Wd*100, color="#5b9bd5", ls="--", alpha=0.4)
    axes[1].fill_between(wd_r*100, [w*100 for w in waccs2], alpha=0.1, color="#5b9bd5")
    axes[1].set_xlabel("Weight of Debt (%)", color="#8b6914", fontsize=8)
    axes[1].set_ylabel("WACC (%)", color="#8b6914", fontsize=8)
    axes[1].set_title("WACC vs Leverage", color="#c8a84b", fontsize=9, fontweight="bold")
    fig.suptitle("Sensitivity Analysis", color="#c8a84b", fontsize=11, fontweight="bold", y=1.03)
    plt.tight_layout(); return fig

def dark_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(BG_DARK); canvas.rect(0,0,W,H,fill=1,stroke=0)
    canvas.setFillColor(GOLD); canvas.rect(0,H-3,W,3,fill=1,stroke=0); canvas.rect(0,0,W,3,fill=1,stroke=0)
    canvas.setFont("Helvetica",7); canvas.setFillColor(GOLD_DARK)
    canvas.drawCentredString(W/2, 10, "WACC = We*Re + Wp*Rp + Wd*Rd*(1-T)   |   AI-Generated Analysis   |   For educational use")
    canvas.restoreState()

def cover_page_callback(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(BG_DARK); canvas.rect(0,0,W,H,fill=1,stroke=0)
    canvas.setFillColor(GOLD); canvas.rect(0,H-4,W,4,fill=1,stroke=0); canvas.rect(0,0,W,4,fill=1,stroke=0)
    canvas.restoreState()

def get_ai_summary(r, api_key):
    api_key = api_key if (api_key and api_key != "hardcoded") else "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"
    try:
        import urllib.request, json as _json
        bench = INDUSTRY_BENCHMARKS.get(r.get("industry", ""), None)
        wacc_pct = r["wacc"] * 100
        company = r.get("company") or "The firm"
        prompt = (
            "You are a senior Wall Street analyst writing the executive summary of a formal Cost of Capital report.\n\n"
            f"Company: {company}\n"
            f"Industry: {r.get('industry', 'Not specified')}\n"
            f"WACC: {wacc_pct:.4f}%\n"
            f"Cost of Equity (Re): {r['Re']*100:.2f}%\n"
            f"Cost of Debt (Rd): {r['Rd']*100:.2f}% (after-tax: {r['after_tax_debt']*100:.2f}%)\n"
            f"Tax Rate: {r['T']*100:.2f}%\n"
            f"Capital Weights: Equity {r['We']*100:.1f}%, Debt {r['Wd']*100:.1f}%, Preferred {r['Wp']*100:.1f}%\n"
            + (f"Industry avg WACC: {bench}%\n" if bench else "")
            + (f"CAPM-implied Re: {r['capm_re']*100:.2f}%\n" if r.get("capm_re") else "")
            + "\nWrite a professional 5-paragraph executive summary. Cover: (1) WACC result and 2025 market context, "
            "(2) capital structure and tax shield, (3) cost of equity vs market, (4) investment hurdle rate implications, "
            "(5) 2-3 specific strategic recommendations. Write like a Goldman Sachs analyst report. "
            "Flowing paragraphs only. No bullet points. No headers."
        )
        req = urllib.request.Request(
            "https://api.openai.com/v1/chat/completions",
            data=_json.dumps({
                "model": "gpt-4o",
                "max_tokens": 1200,
                "messages": [{"role": "user", "content": prompt}]
            }).encode(),
            headers={"Authorization": f"Bearer {api_key}", "content-type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req) as resp:
            data = _json.loads(resp.read())
        return data["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"AI SUMMARY FAILED: {e}")
        return r.get("summary", "")


def generate_wacc_pdf(r, api_key=None):
    buf = io.BytesIO()
    styles = build_styles()
    margin = 20*mm
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=margin, rightMargin=margin,
                             topMargin=margin, bottomMargin=20*mm)
    story = []
    usable_w = W - 2*margin

    # ── Get AI or fallback summary ──
    if api_key:
        exec_summary = get_ai_summary(r, api_key)
    else:
        exec_summary = r.get("summary", "")

    # ── COVER PAGE ──
    story.append(Spacer(1, 40*mm))
    story.append(Paragraph("COST OF CAPITAL", styles["DocTitle"]))
    story.append(Paragraph("Financial Analysis Report", styles["DocSubtitle"]))
    story.append(Spacer(1,5*mm))
    if r.get("company"):
        story.append(Paragraph(r["company"], styles["DocCompany"]))
    if r.get("industry") and r["industry"] != "-- Select --":
        story.append(Paragraph(r["industry"], styles["DateText"]))
    story.append(Spacer(1,12*mm))
    story.append(GoldRule(usable_w))
    story.append(Spacer(1,12*mm))
    hero_data = [
        [Paragraph("WEIGHTED AVERAGE COST OF CAPITAL", styles["WACCLabel"])],
        [Paragraph(f"{r['wacc']*100:.4f}%", styles["WACCHero"])],
    ]
    hero_table = Table(hero_data, colWidths=[usable_w])
    hero_table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),BG_CARD),
        ("BOX",(0,0),(-1,-1),0.5,GOLD),
        ("TOPPADDING",(0,0),(-1,-1),12),
        ("BOTTOMPADDING",(0,0),(-1,-1),12),
    ]))
    story.append(hero_table)
    story.append(Spacer(1,8*mm))

    # AI badge on cover
    ai_badge = Table([[Paragraph(
        "✦ AI-Enhanced Executive Summary" if api_key else "✦ Executive Summary",
        ParagraphStyle("Badge", fontName="Helvetica", fontSize=8,
                       textColor=colors.HexColor("#4caf50" if api_key else "#c8a84b"),
                       alignment=TA_CENTER)
    )]], colWidths=[usable_w])
    ai_badge.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#0a1500" if api_key else "#1a1200")),
        ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#4caf50" if api_key else "#c8a84b")),
        ("TOPPADDING",(0,0),(-1,-1),6), ("BOTTOMPADDING",(0,0),(-1,-1),6),
    ]))
    story.append(ai_badge)
    story.append(Spacer(1,8*mm))
    story.append(GoldRule(usable_w))
    story.append(Spacer(1,8*mm))
    story.append(Paragraph(f"Generated: {r.get('timestamp','')}", styles["DateText"]))
    story.append(PageBreak())

    # ── EXECUTIVE SUMMARY ──
    story.append(Paragraph("EXECUTIVE SUMMARY", styles["SectionTag"]))
    story.append(GoldRule(usable_w, 0.5))
    story.append(Spacer(1,5*mm))
    story.append(Paragraph(r.get("headline",""), styles["Headline"]))
    story.append(Spacer(1,3*mm))

    # Split AI summary into paragraphs
    for para in exec_summary.split(chr(10)+chr(10)):
        para = para.strip()
        if para:
            story.append(Paragraph(para, styles["Body"]))
            story.append(Spacer(1, 3*mm))

    # Industry benchmark callout if available
    bench = INDUSTRY_BENCHMARKS.get(r.get("industry",""), None)
    if bench:
        wacc_pct = r["wacc"]*100
        diff = wacc_pct - bench
        direction = "above" if diff > 0 else "below"
        color = "#ef5350" if diff > 0 else "#4caf50"
        callout_style = ParagraphStyle("Callout", fontName="Helvetica-Bold", fontSize=9,
                                        textColor=colors.HexColor(color),
                                        backColor=colors.HexColor("#0f0c00"),
                                        borderColor=colors.HexColor(color),
                                        borderWidth=1, borderPadding=8,
                                        spaceAfter=6, leading=14)
        story.append(Paragraph(
            f"Industry Benchmark: {r.get('industry','')} sector avg WACC = {bench:.1f}% — "
            f"{r.get('company','Firm')} is {abs(diff):.2f}% {direction} industry average.",
            callout_style
        ))

    story.append(PageBreak())

    # ── CHARTS ──
    story.append(Paragraph("CAPITAL STRUCTURE & WACC ANALYSIS", styles["SectionTag"]))
    story.append(GoldRule(usable_w, 0.5))
    story.append(Spacer(1,5*mm))
    bar_img = fig_to_image(make_bar_chart(r["We"],r["Re"],r["Wd"],r["after_tax_debt"],r["Wp"],r["Rp"],r["wacc"]), usable_w*0.62, 120)
    pie_img = fig_to_image(make_pie_chart(r["We"],r["Wd"],r["Wp"]), usable_w*0.34, 120)
    chart_row = Table([[bar_img, pie_img]], colWidths=[usable_w*0.63, usable_w*0.35])
    chart_row.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),
                                    ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))
    story.append(chart_row)
    story.append(Spacer(1,5*mm))
    sens_img = fig_to_image(make_sensitivity(r["Re"],r["Rd"],r["T"],r["We"],r["Wd"],r["Wp"],r["Rp"]), usable_w, 120)
    story.append(sens_img)
    story.append(PageBreak())

    # ── INPUT PARAMETERS ──
    story.append(Paragraph("INPUT PARAMETERS", styles["SectionTag"]))
    story.append(GoldRule(usable_w, 0.5))
    story.append(Spacer(1,5*mm))
    rows = [
        ["Parameter", "Value", "Notes"],
        ["Cost of Equity (Re)", f"{r['Re']*100:.2f}%", "Required return by equity holders"],
        ["Cost of Debt (Rd)", f"{r['Rd']*100:.2f}%", "Pre-tax rate on debt"],
        ["After-Tax Cost of Debt", f"{r['after_tax_debt']*100:.2f}%", "Rd x (1-T)"],
        ["Cost of Preferred (Rp)", f"{r['Rp']*100:.2f}%", "Preferred dividend yield"],
        ["Corporate Tax Rate (T)", f"{r['T']*100:.2f}%", "Marginal tax rate"],
        ["Weight of Equity (We)", f"{r['We']*100:.2f}%", "Equity % of total capital"],
        ["Weight of Debt (Wd)", f"{r['Wd']*100:.2f}%", "Debt % of total capital"],
        ["Weight of Preferred (Wp)", f"{r['Wp']*100:.2f}%", "Preferred % of total capital"],
        ["WACC", f"{r['wacc']*100:.4f}%", "Weighted Average Cost of Capital"],
    ]
    if r.get("capm_re") is not None:
        rows += [
            ["Risk-Free Rate (Rf)", f"{r['Rf']*100:.2f}%", "10-yr Treasury yield"],
            ["Equity Beta", f"{r['beta']:.2f}", "Systematic risk vs market"],
            ["Market Return (Rm)", f"{r['Rm']*100:.2f}%", "Expected market return"],
            ["CAPM-Implied Re", f"{r['capm_re']*100:.2f}%", "Rf + Beta*(Rm-Rf)"],
        ]
    col_w = [usable_w*0.38, usable_w*0.18, usable_w*0.44]
    t = Table(rows, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),GOLD_DARK),
        ("TEXTCOLOR",(0,0),(-1,0),BG_DARK),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,0),8),
        ("TOPPADDING",(0,0),(-1,-1),5), ("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("FONTNAME",(0,1),(-1,-1),"Helvetica"), ("FONTSIZE",(0,1),(-1,-1),8.5),
        ("TEXTCOLOR",(0,1),(-1,-1),TEXT_BODY),
        ("TEXTCOLOR",(1,1),(1,-1),GOLD_LIGHT), ("FONTNAME",(1,1),(1,-1),"Helvetica-Bold"),
        ("ALIGN",(1,0),(1,-1),"CENTER"), ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LINEBELOW",(0,0),(-1,-1),0.2,colors.HexColor("#c8a84b22")),
        *[("BACKGROUND",(0,i),(-1,i),BG_MID if i%2==1 else BG_DARK) for i in range(1,len(rows))],
        ("BACKGROUND",(0,9),(-1,9),BG_CARD),
        ("TEXTCOLOR",(0,9),(-1,9),GOLD_LIGHT), ("FONTNAME",(0,9),(-1,9),"Helvetica-Bold"),
    ]))
    story.append(t)
    story.append(Spacer(1,8*mm))
    story.append(GoldRule(usable_w, 0.5))
    story.append(Spacer(1,4*mm))
    story.append(Paragraph("WACC = We*Re + Wp*Rp + Wd*Rd*(1-T)",
        ParagraphStyle("Formula", fontName="Helvetica-Bold", fontSize=11,
                       textColor=GOLD, alignment=TA_CENTER, leading=16)))

    doc.build(story, onFirstPage=cover_page_callback, onLaterPages=dark_page)
    buf.seek(0)
    return buf.read()
