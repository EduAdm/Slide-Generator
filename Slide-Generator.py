from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# Initialize 16:9 Widescreen Presentation
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# --- RESTRAINED COLOR PALETTE ---
BG_COLOR = RGBColor(248, 250, 252)       # Light Slate / Off-White
NAVY = RGBColor(15, 23, 42)             # Deep Slate Navy
BRONZE = RGBColor(154, 52, 18)          # Warm Amber Accent
MUTED_TEXT = RGBColor(100, 116, 139)    # Slate 500
CARD_BG = RGBColor(255, 255, 255)       # Pure White Container
BORDER_COLOR = RGBColor(226, 232, 240)  # Border Line

def apply_background(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

def add_header(slide, category, title):
    """Adds header with corrected vertical buffer to prevent line collisions."""
    eyebrow_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.3))
    tf_e = eyebrow_box.text_frame
    tf_e.word_wrap = True
    p_e = tf_e.paragraphs[0]
    p_e.text = category.upper()
    p_e.font.size = Pt(10)
    p_e.font.bold = True
    p_e.font.color.rgb = BRONZE

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.733), Inches(0.7))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = title
    p_t.font.size = Pt(20)
    p_t.font.bold = True
    p_t.font.color.rgb = NAVY

def add_card(slide, left, top, width, height):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = CARD_BG
    shape.line.color.rgb = BORDER_COLOR
    shape.line.width = Pt(1)
    return shape

blank_layout = prs.slide_layouts[6]

# ==========================================
# SLIDE 1: Title Slide
# ==========================================
s1 = prs.slides.add_slide(blank_layout)
apply_background(s1)
add_card(s1, Inches(0.8), Inches(1.2), Inches(11.733), Inches(5.1))

tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(1.6), Inches(10.9), Inches(4.3))
tf1 = tb1.text_frame
tf1.word_wrap = True

p = tf1.paragraphs[0]
p.text = "STEERING COMMITTEE PRESENTATION PLAYBOOK"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = BRONZE
p.space_after = Pt(12)

p = tf1.add_paragraph()
p.text = "AtlasRetail Steering Committee Presentation"
p.font.size = Pt(30)
p.font.bold = True
p.font.color.rgb = NAVY
p.space_after = Pt(10)

p = tf1.add_paragraph()
p.text = "Trade Promotion Optimization (TPO): Commercial Strategy & Growth Engine"
p.font.size = Pt(16)
p.font.color.rgb = MUTED_TEXT
p.space_after = Pt(36)

p = tf1.add_paragraph()
p.text = "OWNER: Jeff K., Sales Manager — ZevationIndustrialUS\nAUDIENCE: Sales Directors, Sales VP, Account Executives | DATA BASIS: May 15, 2026 Baseline"
p.font.size = Pt(11)
p.font.color.rgb = NAVY

# ==========================================
# SLIDE 2: Executive Summary & Context
# ==========================================
s2 = prs.slides.add_slide(blank_layout)
apply_background(s2)
add_header(s2, "Executive Summary & Strategic Context", "AtlasRetail Can Capture $31.1M Net Profit Expansion Without Reducing Top-Line Sales")

cards2 = [
    ("1. Sub-Optimal Spend", "$90M Trade Pool At Risk", [
        "AtlasRetail baseline: $2.0B annual revenue with 18.0% trade spend intensity ($360M pool).",
        "25% of spend ($90M) is sub-optimal, heavily concentrated in low-lift TPRs (1.8 average lift).",
        "High promotional wear-out and retailer forward-buying eroding profit margins."
    ]),
    ("2. Dual-Engine Solution", "Elimination & Reallocation", [
        "Eliminate non-performing spend ($18M pruned at 1.0 lift factor).",
        "Reallocate $36M into high-engagement In-Store Display (3.2 lift) & Feature Circulars (2.4 lift).",
        "Preserve $36M in baseline TPR spend to protect critical retail partner relationships."
    ]),
    ("3. Financial Impact", "Unassailable Return", [
        "$60.7M cumulative 3-year gross profit uplift ($58.1M cumulative net cash value).",
        "Year 1 cash outlay ($1.25M) is self-funding within just 1.9 months of go-live.",
        "Delivers 39.7x steady-state technology ROI vs. AtlasRetail's 5.0x hurdle."
    ])
]

for i, (tag, title, bullets) in enumerate(cards2):
    left = Inches(0.8 + i * 3.98)
    add_card(s2, left, Inches(1.6), Inches(3.78), Inches(5.2))
    tb = s2.shapes.add_textbox(left + Inches(0.2), Inches(1.8), Inches(3.38), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = tag.upper()
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = BRONZE
    p.space_after = Pt(6)
    
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(12)
    
    for b in bullets:
        pb = tf.add_paragraph()
        pb.text = "• " + b
        pb.font.size = Pt(12)
        pb.font.color.rgb = NAVY
        pb.space_after = Pt(8)

# ==========================================
# SLIDE 3: Baseline Assessment & Benchmarks
# ==========================================
s3 = prs.slides.add_slide(blank_layout)
apply_background(s3)
add_header(s3, "Baseline Assessment & Commercial Parameters", "Current $360M Spend is Heavily Skewed Toward Sub-Optimal Price Reductions")

add_card(s3, Inches(0.8), Inches(1.6), Inches(11.733), Inches(5.2))
t_shape3 = s3.shapes.add_table(5, 4, Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.8))
t3 = t_shape3.table
t3.columns[0].width = Inches(3.2)
t3.columns[1].width = Inches(2.0)
t3.columns[2].width = Inches(2.2)
t3.columns[3].width = Inches(3.933)

t3_data = [
    ["Promotion Vehicle Type", "Spend Share (%)", "Annual Spend ($)", "Average Lift Factor & Performance"],
    ["TPR (Temporary Price Reduction)", "55.0%", "$198,000,000", "1.8 Lift — Heavy over-reliance & margin erosion"],
    ["In-Store Display", "22.0%", "$79,200,000", "3.2 Lift — High-impact secondary placement"],
    ["Feature Ad (Circular / Co-Op)", "15.0%", "$54,000,000", "2.4 Lift — Targeted retail circular co-op"],
    ["Digital & Shopper Marketing", "8.0%", "$28,800,000", "1.5 Lift — Sub-scale budget allocation"]
]

for r_idx, row in enumerate(t3_data):
    for c_idx, val in enumerate(row):
        cell = t3.cell(r_idx, c_idx)
        cell.text = val
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY if r_idx == 0 else (CARD_BG if r_idx % 2 == 0 else BG_COLOR)
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(12)
            p.font.bold = (r_idx == 0)
            p.font.color.rgb = RGBColor(255, 255, 255) if r_idx == 0 else NAVY

# ==========================================
# SLIDE 4: Proposed Solution Architecture
# ==========================================
s4 = prs.slides.add_slide(blank_layout)
apply_background(s4)
add_header(s4, "Value Creation Architecture", "Dual-Engine Strategy: Spend Elimination ($18M) and Spend Reallocation ($36M)")

cards4 = [
    ("Step 1: Segmentation", "Sub-Optimal Pool ($90M)", [
        "40.0% ($36M) 'Stay Put': Retained in baseline TPR to protect retail key account relationships.",
        "60.0% ($54M) 'Addressed': Optimized through Zevation Industrial US predictive engine."
    ]),
    ("Step 2: Mechanism 1", "Spend Elimination ($18M)", [
        "33.3% of addressed spend ($17.98M) represents 1.0-lift non-performing promotions.",
        "Pruning 1.0-lift spend generates direct cash savings dropping 100% to gross profit."
    ]),
    ("Step 3: Mechanism 2", "Spend Reallocation ($36M)", [
        "66.7% ($36.02M) upgraded to high-engagement vehicles.",
        "77.0% ($27.73M) shifted to In-Store Display (3.2 lift).",
        "23.0% ($8.28M) shifted to Feature Ads (2.4 lift)."
    ])
]

for i, (tag, title, bullets) in enumerate(cards4):
    left = Inches(0.8 + i * 3.98)
    add_card(s4, left, Inches(1.6), Inches(3.78), Inches(5.2))
    tb = s4.shapes.add_textbox(left + Inches(0.2), Inches(1.8), Inches(3.38), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = tag.upper()
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = BRONZE
    p.space_after = Pt(6)
    
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(12)
    
    for b in bullets:
        pb = tf.add_paragraph()
        pb.text = "• " + b
        pb.font.size = Pt(12)
        pb.font.color.rgb = NAVY
        pb.space_after = Pt(10)

# ==========================================
# SLIDE 5: Mathematical Walk & Proof Points
# ==========================================
s5 = prs.slides.add_slide(blank_layout)
apply_background(s5)
add_header(s5, "Value Creation Mathematical Walk", "Base Scenario Step-by-Step Proof Point to $31,121,366 Gross Profit Expansion")

add_card(s5, Inches(0.8), Inches(1.6), Inches(11.733), Inches(5.2))
t_shape5 = s5.shapes.add_table(6, 3, Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.8))
t5 = t_shape5.table
t5.columns[0].width = Inches(4.0)
t5.columns[1].width = Inches(2.2)
t5.columns[2].width = Inches(5.133)

t5_data = [
    ["Optimization Step / Metric", "Base Scenario Value", "Strategic Accounting Impact"],
    ["1. Eliminated Trade Spend (1.0 Lift)", "$17,982,000", "Direct cash trade spend savings (100% margin drop-through)"],
    ["2. Reallocated Trade Spend", "$36,018,000", "Upgraded from 1.8 TPR baseline to 3.2 Display / 2.4 Feature"],
    ["   – Net Incremental Promo Sales", "+$43,797,888", "Top-line promotional revenue expansion across retail accounts"],
    ["   – Net GP Uplift from Reallocation", "+$13,139,366", "Net bottom-line profit generated at 30.0% GP margin"],
    ["Total Steady-State GP Uplift (Year 3)", "$31,121,366", "+1.6% Net Margin Expansion on $2,000,000,000 Net Revenue"]
]

for r_idx, row in enumerate(t5_data):
    for c_idx, val in enumerate(row):
        cell = t5.cell(r_idx, c_idx)
        cell.text = val
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY if r_idx == 0 else (CARD_BG if r_idx % 2 == 0 else BG_COLOR)
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(11)
            p.font.bold = (r_idx == 0 or r_idx == 5)
            p.font.color.rgb = RGBColor(255, 255, 255) if r_idx == 0 else NAVY

# ==========================================
# SLIDE 6: 3-Year Financial Model & Cash Flow
# ==========================================
s6 = prs.slides.add_slide(blank_layout)
apply_background(s6)
add_header(s6, "3-Year Financial Projections", "Self-Funding Trajectory Generates $58.1M Net Cumulative Cash Benefit")

cards6 = [
    ("Year 1 (25% Ramp)", "Self-Funding Start", [
        "Gross Profit Uplift: $7,780,342",
        "Cash Tech Outlay: $1,250,000 ($650K SaaS + $400K Fee + $200K Analytics)",
        "Net Cash Benefit: +$6,530,342",
        "Payback Horizon: 1.9 Months"
    ]),
    ("Year 2 (70% Ramp)", "Scale & Acceleration", [
        "Gross Profit Uplift: $21,784,956",
        "Cash Tech Outlay: $650,000 (Recurring SaaS License)",
        "Net Cash Benefit: +$21,134,956",
        "Cumulative Cash: +$27,665,298"
    ]),
    ("Year 3 (100% Steady-State)", "Full Run-Rate", [
        "Gross Profit Uplift: $31,121,366",
        "Cash Tech Outlay: $650,000 (Recurring SaaS License)",
        "Net Cash Benefit: +$30,471,366",
        "Cumulative Cash: +$58,136,664"
    ])
]

for i, (tag, title, bullets) in enumerate(cards6):
    left = Inches(0.8 + i * 3.98)
    add_card(s6, left, Inches(1.6), Inches(3.78), Inches(5.2))
    tb = s6.shapes.add_textbox(left + Inches(0.2), Inches(1.8), Inches(3.38), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = tag.upper()
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = BRONZE
    p.space_after = Pt(6)
    
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(12)
    
    for b in bullets:
        pb = tf.add_paragraph()
        pb.text = "• " + b
        pb.font.size = Pt(12)
        pb.font.color.rgb = NAVY
        pb.space_after = Pt(8)

# ==========================================
# SLIDE 7: Commercial Hurdle & ROI Evaluation
# ==========================================
s7 = prs.slides.add_slide(blank_layout)
apply_background(s7)
add_header(s7, "Commercial Hurdle & ROI Evaluation", "Every Financial Dimension Shatters AtlasRetail's 5.0x Tech Investment Hurdle")

grid7 = [
    ("39.7x ROI", "Amortized Technology ROI", "Crushes 5.0x corporate hurdle by +34.7x clearance on $783K/yr amortized tech cost."),
    ("1.9 Months", "Year 1 Cash Payback", "Recovers $1.25M Year 1 cash outlay in under 60 days vs. 12-month standard window."),
    ("3.9 Months", "Full Contract Payback", "Fully recovers total 3-year technology investment ($2.55M) in Year 1 benefit."),
    ("+14.1%", "Corporate Net Profit Lift", "Expands baseline corporate profit from $220.0M to $251.1M at steady state.")
]

for i, (val, title, desc) in enumerate(grid7):
    col = i % 2
    row = i // 2
    left = Inches(0.8 + col * 5.96)
    top = Inches(1.6 + row * 2.68)
    
    add_card(s7, left, top, Inches(5.76), Inches(2.48))
    tb = s7.shapes.add_textbox(left + Inches(0.3), top + Inches(0.2), Inches(5.16), Inches(2.08))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = val
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = BRONZE
    
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(6)
    
    p = tf.add_paragraph()
    p.text = desc
    p.font.size = Pt(12)
    p.font.color.rgb = MUTED_TEXT

# ==========================================
# SLIDE 8: Conservative Downside Sensitivity
# ==========================================
s8 = prs.slides.add_slide(blank_layout)
apply_background(s8)
add_header(s8, "Conservative Scenario & Sensitivity Analysis", "Business Case Remains Highly Profitable Even Under Severe Operational Constraints")

add_card(s8, Inches(0.8), Inches(1.6), Inches(11.733), Inches(5.2))
t_shape8 = s8.shapes.add_table(6, 4, Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.8))
t8 = t_shape8.table
t8.columns[0].width = Inches(3.2)
t8.columns[1].width = Inches(2.0)
t8.columns[2].width = Inches(2.2)
t8.columns[3].width = Inches(3.933)

t8_data = [
    ["Financial Parameter", "Base Scenario", "Conservative Scenario", "Risk Insulation Rationale"],
    ["Sub-Optimal Stay Put Share", "40.0% ($36M)", "60.0% ($54M)", "Leaves 60% of sub-optimal spend untouched"],
    ["Addressed Sub-Optimal Spend", "$54,000,000", "$36,000,000", "Tests 33.3% reduction in optimized spend"],
    ["Year 3 Steady-State Gross Profit", "$31,121,366", "$20,747,578", "Delivers +$20.7M annual uplift in downside case"],
    ["Steady-State Technology ROI", "39.7x", "26.5x", "Exceeds 5.0x hurdle by +21.5x (5.3x clearance)"],
    ["Year 1 Cash Payback Horizon", "1.9 Months", "2.9 Months", "Full cash outlay recovered in under 90 days"]
]

for r_idx, row in enumerate(t8_data):
    for c_idx, val in enumerate(row):
        cell = t8.cell(r_idx, c_idx)
        cell.text = val
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY if r_idx == 0 else (CARD_BG if r_idx % 2 == 0 else BG_COLOR)
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(11)
            p.font.bold = (r_idx == 0 or r_idx == 4)
            p.font.color.rgb = RGBColor(255, 255, 255) if r_idx == 0 else NAVY

# ==========================================
# SLIDE 9: Stakeholder Alignment (Battlecards)
# ==========================================
s9 = prs.slides.add_slide(blank_layout)
apply_background(s9)
add_header(s9, "Steering Committee Stakeholder Framing", "Tailored Battlecards Address Specific Executive Priorities and Biases")

quads9 = [
    ("CFO: John R. (P&L Certainty)", "Financial Rigor & Auditability", "1.9-month cash payback, $58.1M cumulative cash benefit, and 100% grounded in AtlasRetail baseline data."),
    ("CMO: Shveta A. (Brand Elevation)", "Brand Equity & Merchandising", "Replaces blunt 1.8-lift price slashes with $36M shifted to high-visibility In-Store Display (3.2 lift) & Feature Circulars."),
    ("VP RGM: Martin H. (Decision Engine)", "Closed-Loop Governance", "Delivers automated POS reconciliation, baseline vs. incremental attribution, and Year 1 Managed Analytics ($200K)."),
    ("Head of Sales: Anil Y. (Growth)", "Collaborative Retail JBP", "Preserves $36M in TPR spend, generates +$43.8M top-line sales, and equips AEs with JBP co-creation toolsets.")
]

for i, (tag, title, desc) in enumerate(quads9):
    col = i % 2
    row = i // 2
    left = Inches(0.8 + col * 5.96)
    top = Inches(1.6 + row * 2.68)
    
    add_card(s9, left, top, Inches(5.76), Inches(2.48))
    tb = s9.shapes.add_textbox(left + Inches(0.3), top + Inches(0.2), Inches(5.16), Inches(2.08))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = tag.upper()
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = BRONZE
    p.space_after = Pt(4)
    
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(6)
    
    p = tf.add_paragraph()
    p.text = desc
    p.font.size = Pt(12)
    p.font.color.rgb = NAVY

# ==========================================
# SLIDE 10: Implementation Roadmap
# ==========================================
s10 = prs.slides.add_slide(blank_layout)
apply_background(s10)
add_header(s10, "Phased Implementation Roadmap", "Corrected Layout & Column Widths for Scanning Execution Timelines")

add_card(s10, Inches(0.8), Inches(1.6), Inches(11.733), Inches(5.2))
t_shape10 = s10.shapes.add_table(5, 4, Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.8))
t10 = t_shape10.table
# FIXED COLUMN WIDTHS: Prevents single-word vertical wrapping ("Months \n 1 - 3")
t10.columns[0].width = Inches(2.2)
t10.columns[1].width = Inches(1.8)
t10.columns[2].width = Inches(4.333)
t10.columns[3].width = Inches(3.0)

t10_data = [
    ["Phase", "Timeline Window", "Core Operational Deliverables", "Target Financial Impact"],
    ["Phase 1: Foundation", "Months 1 – 3", "Ingest POS/spend data; calibrate SKU models; configure ERP trade accruals.", "Zero baseline disruption; workflow alignment."],
    ["Phase 2: Pilot Go-Live", "Months 4 – 6", "Deploy TPO in 2 lead categories; embed Year 1 Managed Analytics team.", "Early 1.0-lift pruning; initial lift validation."],
    ["Phase 3: Enterprise Rollout", "Months 7 – 18", "Expand across all retail accounts; automate POS deduction reconciliation.", "Achieve Year 1 GP uplift ($7.8M); 1.9mo payback."],
    ["Phase 4: Full Maturity", "Months 19 – 36", "Full automated closed-loop optimization; dynamic scenario modeling.", "Full steady-state GP uplift ($31.1M/yr); +1.6% margin."]
]

for r_idx, row in enumerate(t10_data):
    for c_idx, val in enumerate(row):
        cell = t10.cell(r_idx, c_idx)
        cell.text = val
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY if r_idx == 0 else (CARD_BG if r_idx % 2 == 0 else BG_COLOR)
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(11)
            p.font.bold = (r_idx == 0)
            p.font.color.rgb = RGBColor(255, 255, 255) if r_idx == 0 else NAVY

# ==========================================
# SLIDE 11: Key Objection Handling
# ==========================================
s11 = prs.slides.add_slide(blank_layout)
apply_background(s11)
add_header(s11, "Key Objection Handling Battlecards", "Proactively Mitigating Executive Steering Committee Risks")

cards11 = [
    ("Objection 1: CFO (Risk)", "Implementation Overruns", "Objection: Software implementations promise high ROI but frequently experience cost overruns.\n\nResponse: Year 1 cash outlay ($1.25M) is fully self-funding in 1.9 months. Conservative downside modeling proves 26.5x ROI even under constrained adoption."),
    ("Objection 2: Sales (Relationships)", "Retailer Alienation Risk", "Objection: Cutting trade budgets will alienate key retail buyers and threaten volume quotas.\n\nResponse: We preserve $36M in baseline TPR spend, upgrade $36M to high-lift Display/Feature, and drive +$43.8M in incremental sales for buyers."),
    ("Objection 3: CMO (Brand)", "Generic Benchmarks", "Objection: Generic benchmarks don't reflect our premium brand equity across diverse categories.\n\nResponse: We replace blunt price slashes with SKU-level elasticity models calibrated to AtlasRetail's specific categories.")
]

for i, (tag, title, desc) in enumerate(cards11):
    left = Inches(0.8 + i * 3.98)
    add_card(s11, left, Inches(1.6), Inches(3.78), Inches(5.2))
    tb = s11.shapes.add_textbox(left + Inches(0.2), Inches(1.8), Inches(3.38), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = tag.upper()
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = BRONZE
    p.space_after = Pt(6)
    
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(12)
    
    p = tf.add_paragraph()
    p.text = desc
    p.font.size = Pt(12)
    p.font.color.rgb = NAVY

# ==========================================
# SLIDE 12: Pre-Meeting Action Plan
# ==========================================
s12 = prs.slides.add_slide(blank_layout)
apply_background(s12)
add_header(s12, "Immediate Pre-Meeting Action Items", "60-Minute Choreography & Immediate Next Steps to Win Mandate")

add_card(s12, Inches(0.8), Inches(1.6), Inches(7.5), Inches(5.2))
tb12_l = s12.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(7.1), Inches(4.8))
tf12_l = tb12_l.text_frame
tf12_l.word_wrap = True

p = tf12_l.paragraphs[0]
p.text = "IMMEDIATE PRE-MEETING ACTION ITEMS"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = BRONZE
p.space_after = Pt(12)

actions12 = [
    ("1. Pre-Meeting Champion Alignment", "Conduct 1-on-1 prep session with Martin H. (VP RGM) to pre-align on platform demo workflows and baseline attribution models."),
    ("2. Sales Head De-escalation", "Hold informal pre-briefing with Anil Y. (Head of Sales) to demonstrate AE JBP scorecard and reassure him no 'waste' rhetoric will be used."),
    ("3. CFO Financial Tear-Sheet", "Prepare one-page executive financial tear-sheet for John R. (CFO) summarizing the 1.9-month payback and 39.7x steady ROI.")
]

for title, desc in actions12:
    p = tf12_l.add_paragraph()
    p.text = title
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = NAVY
    
    p = tf12_l.add_paragraph()
    p.text = desc
    p.font.size = Pt(12)
    p.font.color.rgb = MUTED_TEXT
    p.space_after = Pt(12)

add_card(s12, Inches(8.6), Inches(1.6), Inches(3.933), Inches(5.2))
tb12_r = s12.shapes.add_textbox(Inches(8.8), Inches(1.8), Inches(3.533), Inches(4.8))
tf12_r = tb12_r.text_frame
tf12_r.word_wrap = True

p = tf12_r.paragraphs[0]
p.text = "WINNING THE MANDATE"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = BRONZE
p.space_after = Pt(14)

p = tf12_r.add_paragraph()
p.text = "Zevation Industrial US is perfectly positioned to win the AtlasRetail Steering Committee. We have the data, the financial walk, the commercial proof points, and the stakeholder strategy to turn promotional challenges into an enterprise growth engine."
p.font.size = Pt(13)
p.font.color.rgb = NAVY
p.space_after = Pt(20)

p = tf12_r.add_paragraph()
p.text = "$31.1M Net Profit Lift\n39.7x Tech ROI\n1.9 Month Cash Payback"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = BRONZE

# Save Complete Presentation
prs.save("Executive_Steering_Committee_Deck.pptx")
print("Complete 12-slide presentation successfully generated: Executive_Steering_Committee_Deck.pptx")
