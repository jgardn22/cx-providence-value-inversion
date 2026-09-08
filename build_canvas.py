# -*- coding: utf-8 -*-
"""
Builds raw/decision-canvas.html: the Providence "Enterprise Application
Value Inversion" interactive decision canvas.

Format: follows the same click-driven tab/panel architecture as
collab-hubs/kaiser-risant-value-inversion/raw/decision-canvas.html and
collab-hubs/centene-value-inversion/raw/decision-canvas.html (Problem /
Status Quo / The Big Question / Opportunity / New Approach / Plan),
reskinned with Providence's brand colors (pulled live from providence.org
9/8/2026) and populated with Providence-specific figures.

Providence is a provider health system (not a payer like Centene), so the
Status Quo and Proof-it-works content leans toward the original clinical
care-journey framing rather than Centene's payer-specific one.

Self-contained single HTML file: logos are embedded as base64 data URIs so
the file has zero external asset dependencies (works after passcode
encryption wraps the whole document).
"""
import json
import os

BASE = os.path.dirname(__file__)
ASSETS = os.path.join(BASE, "assets")
RAW_DIR = os.path.join(BASE, "raw")
os.makedirs(RAW_DIR, exist_ok=True)

with open(os.path.join(ASSETS, "logos-b64.json"), "r") as f:
    LOGOS = json.load(f)

PROVIDENCE_COLOR = f"data:image/png;base64,{LOGOS['providence_color']}"
PROVIDENCE_WHITE = f"data:image/png;base64,{LOGOS['providence_white']}"

# Deck-style screenshots reused as-is from the Kaiser Permanente + Risant
# Health build: these are generic Microsoft narrative/architecture graphics
# with no customer-specific branding or numbers baked into the pixels (the
# Blind Spot diagram, the full Microsoft IQ platform architecture, the
# Enterprise Data Unlocked strip, and the Big Question pause-moment slide).
IMG_BLIND_SPOT = f"data:image/png;base64,{LOGOS['blind_spot']}"
IMG_IQ_PLATFORM = f"data:image/png;base64,{LOGOS['iq_platform']}"
IMG_ENTERPRISE_DATA = f"data:image/png;base64,{LOGOS['enterprise_data']}"
IMG_BIG_QUESTION = f"data:image/png;base64,{LOGOS['big_question']}"

# ============================================================ MODEL =======
# Providence "Estimated Software Spend" TAM (xlsx supplied by Jake, read via
# Excel COM automation since the file is IRM-protected; TPID PROVIDENCE ST.
# JOSEPH HEALTH - 1962997, HQ scope). ERP excludes Manufacturing Product Life
# Cycle Management ($1.56M): no credible Dynamics 365 replacement exists for
# it, per the prior Microsoft/Providence TAM analysis (Providence-AIBS-TAM
# slides 5-6), which independently confirms this same exclusion methodology
# and arrives at the same ~$252M total.
TAM_CATEGORY_TOTALS = {
    "App Development": 23.702001,
    "CRM": 85.969812,
    "ERP": 143.993573 - 1.556262,
}
TAM_TOTAL = sum(TAM_CATEGORY_TOTALS.values())  # ~$252.1M

# Illustrative vendor-name examples per category. Genesys is confirmed
# present in Providence's real environment (contact center); the rest are
# illustrative examples of plausible Power Apps / Dynamics 365 replacement
# candidates for a health system of this size and shape.
CATEGORY_VENDOR_EXAMPLES = {
    "App Development": ["ServiceNow", "UiPath", "Automation Anywhere", "Blue Prism"],
    "CRM": ["Salesforce", "Genesys", "Marketo", "NICE", "Avaya", "Cisco"],
    "ERP": ["SAP", "Oracle", "Workday", "Kronos", "Infor", "Concur"],
}

PROVIDENCE_EMPLOYEES = 125_000  # Jake-supplied: "approximately 125,000 caregivers and staff"
POWERAPPS_RATE = 1.85           # Jake-supplied, per user/month (Power Apps Premium)
D365_RATE = 33.44               # Jake-supplied, per user/month

NAPKIN_FLOOR_MONTHLY = POWERAPPS_RATE                # floor: Power Apps only
NAPKIN_CEILING_MONTHLY = POWERAPPS_RATE + D365_RATE  # ceiling: Power Apps + Dynamics 365


def floor_ceiling_m(headcount):
    floor_m = (headcount * NAPKIN_FLOOR_MONTHLY * 12) / 1_000_000
    ceiling_m = (headcount * NAPKIN_CEILING_MONTHLY * 12) / 1_000_000
    return floor_m, ceiling_m


FLOOR_M, CEILING_M = floor_ceiling_m(PROVIDENCE_EMPLOYEES)
FLOOR_SAVINGS_POTENTIAL = TAM_TOTAL - FLOOR_M      # full run-rate $/yr saved, narrow-licensed
CEILING_SAVINGS_POTENTIAL = TAM_TOTAL - CEILING_M  # full run-rate $/yr saved, broadly-licensed

TENYR_YEARS = list(range(1, 11))
TENYR_FLOOR_ANNUAL = [FLOOR_SAVINGS_POTENTIAL * n / 10 for n in TENYR_YEARS]
TENYR_CEILING_ANNUAL = [CEILING_SAVINGS_POTENTIAL * n / 10 for n in TENYR_YEARS]
TENYR_FLOOR_CUMULATIVE = sum(TENYR_FLOOR_ANNUAL)
TENYR_CEILING_CUMULATIVE = sum(TENYR_CEILING_ANNUAL)

print(f"TAM_TOTAL=${TAM_TOTAL:.1f}M  floor_m=${FLOOR_M:.2f}M  ceiling_m=${CEILING_M:.2f}M")
print(f"Yr10 floor savings=${FLOOR_SAVINGS_POTENTIAL:.0f}M/yr  ceiling savings=${CEILING_SAVINGS_POTENTIAL:.0f}M/yr")
print(f"10-yr cumulative: floor=${TENYR_FLOOR_CUMULATIVE:,.0f}M  ceiling=${TENYR_CEILING_CUMULATIVE:,.0f}M")


def fmt_big_m(value_m):
    """Format a $-millions figure, switching to $-billions above 1000."""
    if value_m >= 1000:
        return f"${value_m / 1000:.1f}B"
    return f"${value_m:,.0f}M"


# --- Compounding-return chart (SVG, computed server-side from the real
# Providence model). Total cost of ownership = the remaining, un-migrated
# portion of the TAM footprint each year as a phased 10%/year migration
# progresses (narrow-licensed / floor scenario). Value realization is
# illustrative and unitless, shaped only to cross the TCO curve visually.
_TCO_YEARS = list(range(0, 11))
TCO_BY_YEAR = [TAM_TOTAL - FLOOR_SAVINGS_POTENTIAL * n / 10 for n in _TCO_YEARS]
VALUE_BY_YEAR = [TAM_TOTAL * 1.05 * (n / 10) ** 1.6 for n in _TCO_YEARS]
_CHART_MAX = max(max(TCO_BY_YEAR), max(VALUE_BY_YEAR)) * 1.05
_CHART_X0, _CHART_X1 = 50, 750
_CHART_Y0, _CHART_Y1 = 20, 220  # y0 = top (max value), y1 = bottom (zero)


def _to_svg_points(vals):
    n = len(vals) - 1
    pts = []
    for i, v in enumerate(vals):
        x = _CHART_X0 + (_CHART_X1 - _CHART_X0) * i / n
        y = _CHART_Y1 - (_CHART_Y1 - _CHART_Y0) * (v / _CHART_MAX)
        pts.append(f"{x:.1f},{y:.1f}")
    return " ".join(pts)


TCO_SVG_POINTS = _to_svg_points(TCO_BY_YEAR)
VALUE_SVG_POINTS = _to_svg_points(VALUE_BY_YEAR)
TCO_START_LABEL = fmt_big_m(TCO_BY_YEAR[0])
TCO_END_LABEL = f"${TCO_BY_YEAR[-1]:.1f}M/yr"

# ============================================================== CSS =======
CSS = """
:root {
  color-scheme: light;
  --cp-bg: #f5f7fb;
  --cp-bg-elevated: #eef2f8;
  --cp-surface: #ffffff;
  --cp-surface-soft: #eef2f8;
  --cp-border: #d7dde8;
  --cp-border-strong: #8496b0;
  --cp-text: #16233a;
  --cp-text-muted: #4d5e70;
  --cp-text-soft: #6b7c8f;
  --cp-accent: #00338E;
  --cp-accent-hover: #00266B;
  --cp-accent-soft: rgba(0, 51, 142, 0.10);
  --cp-accent-fg: #ffffff;
  --cp-success: #39892F;
  --cp-danger: #A14545;
  --cp-warning: #f4a100;
  --cp-link: #00338E;
  --cp-shadow: 0 18px 48px rgba(10, 37, 64, 0.14);
  --cp-overlay: rgba(255, 255, 255, 0.82);
  --cp-panel: rgba(255, 255, 255, 0.88);
  --cp-panel-strong: rgba(255, 255, 255, 0.97);
  --cp-sheen: rgba(255, 255, 255, 0.58);
  --cp-highlight: rgba(57, 137, 47, 0.16);
  --cp-purple: #5B3A8E;
  --cp-green: #39892F;
  --cp-teal: #008575;
  --cp-gold: #C9862E;
  --cp-navy-fixed: #071835;
  --cp-navy-fixed-2: #0f2951;
}
html[data-theme="dark"] {
  color-scheme: dark;
  --cp-bg: #071835;
  --cp-bg-elevated: #04122a;
  --cp-surface: #0f2951;
  --cp-surface-soft: #163765;
  --cp-border: #234a80;
  --cp-border-strong: #3d689f;
  --cp-text: #e8eef5;
  --cp-text-muted: #b7c6d9;
  --cp-text-soft: #c9d6e4;
  --cp-accent: #5CB85C;
  --cp-accent-hover: #7ecb7e;
  --cp-accent-soft: rgba(92, 184, 92, 0.16);
  --cp-accent-fg: #071835;
  --cp-success: #5CB85C;
  --cp-danger: #ef8b8b;
  --cp-warning: #ffc44d;
  --cp-link: #5CB85C;
  --cp-shadow: 0 18px 48px rgba(0, 0, 0, 0.32);
  --cp-overlay: rgba(7, 24, 53, 0.88);
  --cp-panel: rgba(15, 41, 81, 0.74);
  --cp-panel-strong: rgba(15, 41, 81, 0.97);
  --cp-sheen: rgba(255, 255, 255, 0.05);
  --cp-highlight: rgba(92, 184, 92, 0.16);
  --cp-purple: #a98fd1;
  --cp-green: #5CB85C;
  --cp-teal: #5fd8c4;
  --cp-gold: #f0b860;
}
* { box-sizing: border-box; }
html { min-height: 100%; }
body {
  margin: 0; font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
  background: var(--cp-bg); color: var(--cp-text); line-height: 1.45;
  transition: background 200ms ease, color 200ms ease;
}
.app { position: relative; z-index: 2; max-width: 1440px; margin: 0 auto; padding: 24px; }

.hero {
  background: var(--cp-panel-strong); border: 1px solid var(--cp-border); border-radius: 20px;
  padding: 32px 36px; margin-bottom: 20px; box-shadow: var(--cp-shadow);
}
.hero-logos { display: flex; align-items: center; gap: 14px; margin-bottom: 18px; }
.hero-logos img { height: 40px; width: auto; }
.eyebrow { display: flex; align-items: center; gap: 8px; margin: 0 0 10px; font-size: 0.82rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--cp-accent); }
.pulse-dot { width: 8px; height: 8px; border-radius: 999px; background: var(--cp-accent); display: inline-block; animation: pulse 2.2s ease infinite; }
@keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.35; } }
h1 { max-width: 980px; margin: 0; font-size: clamp(2rem, 4vw, 4rem); line-height: 1.02; letter-spacing: -0.04em; }
h1 em { font-style: italic; color: var(--cp-accent); white-space: nowrap; }
.hero p.lede { max-width: 900px; margin: 14px 0 0; color: var(--cp-text-muted); font-size: clamp(1rem, 1.5vw, 1.2rem); line-height: 1.5; }
.hero-foot { display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px; margin-top: 22px; font-size: 0.8rem; color: var(--cp-text-soft); }

.tabs {
  display: grid; grid-template-columns: repeat(6, 1fr); gap: 10px; margin-bottom: 20px;
  position: sticky; top: 12px; z-index: 30; padding: 8px; border-radius: 16px;
  background: var(--cp-panel-strong); box-shadow: var(--cp-shadow); backdrop-filter: blur(14px);
}
.tab-button {
  position: relative; min-height: 58px; border: 1px solid var(--cp-border); border-radius: 0.625rem;
  background: var(--cp-surface); color: var(--cp-text); cursor: pointer;
  transition: transform 180ms ease, border-color 180ms ease, background 180ms ease;
  font-weight: 600; font-size: 0.95rem;
}
.tab-button:hover, .tab-button:focus-visible { transform: translateY(-2px); border-color: var(--cp-accent); outline: none; }
.tab-button.active { background: var(--cp-accent); color: var(--cp-accent-fg); border-color: var(--cp-accent); }
.tab-num { display: block; font-size: 0.68rem; opacity: 0.7; font-weight: 700; letter-spacing: 0.06em; margin-bottom: 2px; }

.panel { display: none; min-height: 560px; animation: enter 360ms ease both; }
.panel.active { display: block; }
@keyframes enter { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }

.panel-head { margin: 6px 4px 18px; }
.panel-head h2 { margin: 0 0 6px; font-size: clamp(1.4rem, 2.4vw, 2rem); letter-spacing: -0.02em; }
.panel-head p { margin: 0; color: var(--cp-text-muted); font-size: 1rem; max-width: 900px; }

.card {
  border: 1px solid var(--cp-border); border-radius: 16px; background: var(--cp-panel-strong);
  box-shadow: var(--cp-shadow); padding: 24px; margin-bottom: 20px;
}
.card-header { margin-bottom: 16px; }
.card-header h3 { margin: 0 0 6px; font-size: 1.25rem; }
.card-header p { margin: 0; color: var(--cp-text-muted); font-size: 0.92rem; max-width: 820px; }
.section-grid { display: grid; grid-template-columns: minmax(0, 1.6fr) minmax(260px, 1fr); gap: 20px; align-items: start; }

.twin-cards { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin: 4px 0 18px; }
.twin-card { border-radius: 14px; padding: 18px; border: 1px solid var(--cp-border); }
.twin-card.disruption { background: var(--cp-surface-soft); }
.twin-card.squeeze { background: color-mix(in srgb, var(--cp-danger) 8%, var(--cp-surface)); }
.twin-tag { display: inline-block; padding: 5px 14px; border-radius: 999px; font-size: 0.78rem; font-weight: 700; color: #fff; margin-bottom: 10px; }
.twin-tag.disruption { background: var(--cp-accent); }
.twin-tag.squeeze { background: #A14545; }
.twin-card p { margin: 0; font-size: 0.92rem; line-height: 1.55; color: var(--cp-text); }
.twin-card b { color: var(--cp-accent); }
.twin-card.squeeze b { color: var(--cp-danger); }

.footer-strip {
  margin-top: 4px; padding: 14px 18px; border-radius: 12px; background: var(--cp-accent-soft);
  border: 1px solid var(--cp-border); color: var(--cp-text); font-style: italic; font-size: 0.95rem;
}

.detail-panel {
  border: 1px solid var(--cp-border); border-radius: 14px; background: var(--cp-surface-soft);
  padding: 20px; min-height: 220px; position: sticky; top: 90px;
}
.detail-panel .kicker { margin: 0 0 6px; font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--cp-accent); }
.detail-panel h4 { margin: 0 0 8px; font-size: 1.05rem; }
.detail-panel p { margin: 0 0 8px; color: var(--cp-text-muted); font-size: 0.92rem; line-height: 1.5; }
.detail-panel .tag-row { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 10px; }
.detail-tag { font-size: 0.72rem; padding: 3px 9px; border-radius: 999px; background: var(--cp-accent-soft); color: var(--cp-accent); font-weight: 600; }

/* detail panel that sits centered below a picker row instead of beside it */
.detail-panel-below {
  border: 1px solid var(--cp-border); border-radius: 14px; background: var(--cp-surface-soft);
  padding: 22px; margin-top: 14px; position: static; top: auto; min-height: 0;
}
.detail-panel-below .kicker { margin: 0 0 6px; font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--cp-accent); }
.detail-panel-below h4 { margin: 0 0 8px; font-size: 1.1rem; }
.detail-panel-below p { margin: 0; color: var(--cp-text-muted); font-size: 0.95rem; line-height: 1.55; max-width: 760px; }

/* embedded deck screenshots (source-of-truth imagery we chose not to recreate in CSS) */
.deck-screenshot {
  display: block; width: 100%; height: auto; border-radius: 12px;
  border: 1px solid var(--cp-border); margin: 10px 0;
}

/* full-bleed single-image panel (e.g. The Big Question pause moment): no card,
   no header, no footer strip, just the slide, centered and breathing */
.big-question-panel { display: flex; align-items: center; justify-content: center; min-height: 560px; padding: 20px 0; }
.big-question-panel img {
  display: block; width: 100%; max-width: 1100px; height: auto; border-radius: 18px;
  box-shadow: 0 24px 60px rgba(0,0,0,0.35); border: 1px solid var(--cp-border);
}

/* visually hidden but still exposed to screen readers / accessibility tree */
.sr-only {
  position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px;
  overflow: hidden; clip: rect(0,0,0,0); white-space: nowrap; border: 0;
}

/* "what it takes to get there" numbered step list */
.get-there-list { display: flex; flex-direction: column; gap: 4px; margin-top: 4px; }
.get-there-step { display: flex; align-items: flex-start; gap: 16px; padding: 12px 4px; position: relative; }
.get-there-step:not(:last-child)::after {
  content: ""; position: absolute; left: 19px; top: 46px; bottom: -4px; width: 2px; background: var(--cp-border);
}
.get-there-badge {
  flex-shrink: 0; width: 40px; height: 40px; border-radius: 999px; display: flex; align-items: center;
  justify-content: center; font-weight: 800; font-size: 1.05rem; color: #fff; z-index: 1;
}
.get-there-step h5 { margin: 0 0 4px; font-size: 1rem; }
.get-there-step p { margin: 0; font-size: 0.88rem; color: var(--cp-text-muted); line-height: 1.45; }

/* stat card grid */
.stat-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin-bottom: 6px; }
.stat-card { border: 1px solid var(--cp-border); border-radius: 12px; padding: 18px; background: var(--cp-surface); border-top: 4px solid var(--cp-accent); }
.stat-card .stat-num { font-size: 2rem; font-weight: 800; font-family: inherit; margin-bottom: 6px; }
.stat-card p { margin: 0; font-size: 0.86rem; color: var(--cp-text-muted); line-height: 1.4; }

/* clickable pick group (generic) */
.pick-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 6px; }
.pick-grid-wide { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 6px; }
.pick-btn { display: flex; flex-direction: column; align-items: flex-start; gap: 6px; text-align: left; border: 1px solid var(--cp-border); border-radius: 12px; padding: 14px; background: var(--cp-surface); cursor: pointer; transition: border-color 150ms ease, transform 150ms ease; }
.pick-btn:hover { transform: translateY(-2px); }
.pick-btn.active { border-color: var(--cp-accent); background: var(--cp-accent-soft); }
.pick-index { display: inline-flex; align-items: center; justify-content: center; width: 30px; height: 30px; border-radius: 999px; color: #fff; font-weight: 700; font-size: 0.85rem; }
.pick-title { font-weight: 700; font-size: 0.95rem; }
.pick-meta { font-size: 0.78rem; color: var(--cp-text-soft); }

/* handoff cards (fragmented context) */
.handoff-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-top: 4px; }
.handoff-card { border: 1px solid var(--cp-border); border-top: 4px solid var(--cp-danger); border-radius: 10px; padding: 12px; background: var(--cp-surface); font-size: 0.82rem; }
.handoff-card b { display: block; margin-bottom: 4px; }
.handoff-card .x { color: var(--cp-danger); font-weight: 700; }

/* value stack (Then/Now) */
.stack-wrap { display: grid; grid-template-columns: 1fr auto 1fr; gap: 14px; align-items: center; }
.stack-col-label { display: flex; justify-content: space-between; align-items: center; font-size: 0.78rem; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; color: var(--cp-text-soft); margin-bottom: 8px; }
.stack-col-label .dir-down { color: var(--cp-danger); }
.stack-col-label .dir-up { color: var(--cp-green); }
.stack { display: flex; flex-direction: column; border-radius: 12px; overflow: hidden; height: 320px; }
.stack-seg { display: flex; align-items: center; padding: 0 14px; font-weight: 700; font-size: 0.88rem; cursor: pointer; transition: filter 150ms ease; border: none; text-align: left; }
.stack-seg:hover, .stack-seg.active { filter: brightness(1.08); }
.stack-mid-label { writing-mode: vertical-rl; text-align: center; font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--cp-text-soft); }

/* hyperscaler compare rows */
.compare-row { display: flex; flex-direction: column; gap: 6px; border: 1px solid var(--cp-border); border-radius: 12px; padding: 14px; margin-bottom: 10px; background: var(--cp-surface); cursor: pointer; transition: border-color 150ms ease; }
.compare-row:hover, .compare-row.active { border-color: var(--cp-accent); }
.compare-row .row-label { font-size: 0.72rem; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; color: var(--cp-text-soft); }
.compare-row .row-title { font-weight: 700; font-size: 1rem; }
.compare-row .row-sub { font-size: 0.85rem; color: var(--cp-text-muted); }
.bill-bar { display: flex; height: 16px; border-radius: 999px; overflow: hidden; margin-top: 8px; font-size: 0.62rem; font-weight: 700; color: #fff; }
.bill-bar span { display: flex; align-items: center; justify-content: center; }

/* TAM whitespace bar */
.tam-bar-wrap { margin-bottom: 16px; }
.tam-bar { display: flex; height: 22px; border-radius: 999px; overflow: hidden; font-size: 0.7rem; font-weight: 700; color: #fff; }
.tam-bar .captured { background: var(--cp-border-strong); }
.tam-bar .whitespace { background: var(--cp-accent); display: flex; align-items: center; justify-content: center; }

/* savings bar chart */
.chart-wrap { margin: 10px 0; }
.chart-bars { display: flex; align-items: flex-end; gap: 10px; height: 260px; padding: 0 4px; }
.chart-year { flex: 1; display: flex; flex-direction: column; align-items: center; gap: 4px; height: 100%; justify-content: flex-end; }
.chart-bar-pair { display: flex; align-items: flex-end; gap: 3px; width: 100%; justify-content: center; height: 100%; }
.chart-bar { width: 40%; border-radius: 4px 4px 0 0; position: relative; }
.chart-bar.narrow { background: var(--cp-green); }
.chart-bar.broad { background: var(--cp-accent); }
.chart-bar .bar-label { position: absolute; top: -20px; left: 50%; transform: translateX(-50%); font-size: 0.62rem; font-weight: 700; white-space: nowrap; }
.chart-year-label { font-size: 0.72rem; color: var(--cp-text-soft); margin-top: 4px; }
.chart-legend { display: flex; gap: 18px; margin-top: 14px; flex-wrap: wrap; }
.chart-legend span { display: flex; align-items: center; gap: 6px; font-size: 0.82rem; color: var(--cp-text-muted); }
.chart-legend i { width: 12px; height: 12px; border-radius: 3px; display: inline-block; }

/* compounding-return SVG line chart */
.compound-chart-wrap { margin: 10px 0; }
.compound-chart-wrap svg { width: 100%; height: auto; display: block; }
.compound-chart-legend { display: flex; gap: 20px; margin-top: 8px; flex-wrap: wrap; font-size: 0.82rem; color: var(--cp-text-muted); }
.compound-chart-legend span { display: flex; align-items: center; gap: 6px; }
.compound-chart-legend i { width: 14px; height: 3px; border-radius: 2px; display: inline-block; }

.source-note { font-size: 0.76rem; color: var(--cp-text-soft); font-style: italic; margin-top: 10px; line-height: 1.5; }

.tooltip {
  position: fixed; z-index: 40; max-width: 260px; padding: 10px 14px; border-radius: 10px;
  background: var(--cp-panel-strong); border: 1px solid var(--cp-border); box-shadow: var(--cp-shadow);
  font-size: 0.82rem; color: var(--cp-text); pointer-events: none; opacity: 0; transition: opacity 120ms ease;
}
.tooltip.visible { opacity: 1; }
.tooltip .tip-title { font-weight: 700; margin-bottom: 4px; color: var(--cp-accent); }

/* closing statement + reframe (Plan tab) */
.reframe-grid { display: grid; grid-template-columns: 1fr auto 1fr; gap: 16px; align-items: stretch; }
.reframe-box { border-radius: 14px; padding: 20px; }
.reframe-box.scrutiny { background: color-mix(in srgb, var(--cp-green) 12%, var(--cp-surface)); border: 1px solid var(--cp-border); }
.reframe-box.matters { background: var(--cp-navy-fixed); color: #fff; }
.reframe-tag { display: inline-block; font-size: 0.68rem; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; padding: 3px 10px; border-radius: 999px; }
.reframe-box.scrutiny .reframe-tag { background: var(--cp-green); color: #fff; }
.reframe-box.matters .reframe-tag { background: var(--cp-gold); color: #16233a; }
.reframe-num { font-size: 2.2rem; font-weight: 800; margin-top: 8px; }
.reframe-arrow { display: flex; align-items: center; justify-content: center; font-size: 1.4rem; color: var(--cp-text-soft); }

.close-card {
  background: var(--cp-navy-fixed); background-image: linear-gradient(135deg, var(--cp-navy-fixed), var(--cp-navy-fixed-2));
  color: #fff; border-radius: 20px; padding: 40px 32px; text-align: center; margin-top: 20px;
}
.close-logos { display: flex; align-items: center; justify-content: center; gap: 14px; margin-bottom: 20px; }
.close-logos img { height: 28px; }
.close-card h3 { font-size: clamp(1.3rem, 2.6vw, 1.9rem); max-width: 900px; margin: 0 auto 8px; line-height: 1.35; }
.close-card h3 .muted { opacity: 0.55; }
.close-card h3 .hi { color: var(--cp-accent); }
.next-step-box { max-width: 620px; margin: 22px auto 0; padding: 16px 20px; border: 1px solid rgba(255,255,255,0.25); border-radius: 12px; background: rgba(255,255,255,0.06); }
.next-step-box .tag { font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; color: var(--cp-gold); margin-bottom: 6px; }
.next-step-box p { margin: 0; font-size: 0.92rem; opacity: 0.92; }

footer.hub-footer { text-align: center; padding: 24px; color: var(--cp-text-soft); font-size: 0.8rem; }

@media (max-width: 1060px) {
  .section-grid, .stack-wrap, .reframe-grid, .twin-cards { grid-template-columns: 1fr; }
  .handoff-row { grid-template-columns: repeat(2, 1fr); }
  .tabs { grid-template-columns: repeat(3, 1fr); }
  .detail-panel { position: static; }
  .stack { height: auto; }
  .stack-seg { min-height: 46px; }
  .stack-mid-label { writing-mode: horizontal-tb; padding: 6px 0; }
  .reframe-arrow { transform: rotate(90deg); padding: 6px 0; }
}
@media (max-width: 640px) {
  .app { padding: 14px; }
  .hero { padding: 22px; }
  .tabs { grid-template-columns: repeat(2, 1fr); position: static; }
  .chart-bars { height: 200px; }
}
"""
print("CSS length:", len(CSS))

# ============================================================= BODY =======
HERO = f"""
<section class="hero">
  <div class="hero-logos">
    <img id="logo-providence" src="{PROVIDENCE_COLOR}" alt="Providence" />
  </div>
  <p class="eyebrow"><span class="pulse-dot" aria-hidden="true"></span> Providence &middot; Executive decision canvas</p>
  <h1>The Enterprise Application<br/><em>Value Inversion</em></h1>
  <p class="lede">The total cost of your application portfolio is inverting. This canvas walks the strategic argument, Providence-specific economics, and the practical next step: click through Problem, Status Quo, The Big Question, Opportunity, New Approach, and Plan.</p>
  <div class="hero-foot">
    <span>Prepared by Microsoft for Providence &middot; confidential</span>
    <span>All figures illustrative unless noted; see Plan tab for sourcing</span>
  </div>
</section>
"""

TABS_NAV = """
<nav class="tabs" aria-label="Decision canvas sections">
  <button class="tab-button active" data-tab="problem" data-tip-title="Problem" data-tip="Set up why application costs keep climbing: the SaaSpocalypse, the value inversion, the hyperscaler tax, and the compounding cost of bolted-on AI.">
    <span class="tab-num">01</span>Problem
  </button>
  <button class="tab-button" data-tab="status-quo" data-tip-title="Status Quo" data-tip="Show what the status quo approach actually looks like today: Copilots that only see one system, and the fragmented-context research proving why most enterprise AI never pays off.">
    <span class="tab-num">02</span>Status Quo
  </button>
  <button class="tab-button" data-tab="big-question" data-tip-title="The Big Question" data-tip="The burner is already on. A pause moment: in five years, will you look back and wish you'd taken a different strategy?">
    <span class="tab-num">03</span>The Big Question
  </button>
  <button class="tab-button" data-tab="opportunity" data-tip-title="Opportunity" data-tip="Ground the argument in Providence's own numbers: the $252.1M TAM and the 10-year savings model.">
    <span class="tab-num">04</span>Opportunity
  </button>
  <button class="tab-button" data-tab="new-approach" data-tip-title="New Approach" data-tip="The blue ocean strategy: one platform with five connected IQs, the full Microsoft IQ architecture, and proof the compounding return is real.">
    <span class="tab-num">05</span>New Approach
  </button>
  <button class="tab-button" data-tab="plan" data-tip-title="Plan" data-tip="The harder question that matters more than scrutiny, and what it takes to get there.">
    <span class="tab-num">06</span>Plan
  </button>
</nav>
"""

PANEL_PROBLEM = """
<section id="problem" class="panel active" aria-labelledby="problem-title">
  <div class="panel-head">
    <h2 id="problem-title">Costs are climbing because value already moved.</h2>
    <p>AI didn't add a feature to enterprise software. It changed where value lives, and that shift is what's driving cost up across the portfolio.</p>
  </div>

  <div class="twin-cards">
    <div class="twin-card disruption">
      <span class="twin-tag disruption">The disruption</span>
      <p>Agentic AI can now perform the work SaaS interfaces were built for. No human has to open the tool. <b>Legacy SaaS valuations</b> already took real hits in 2026 as capital rotates toward AI infrastructure. Some call it the <b>&ldquo;SaaSpocalypse.&rdquo;</b></p>
    </div>
    <div class="twin-card squeeze">
      <span class="twin-tag squeeze">The cost squeeze</span>
      <p>At the same time, the SaaS you already own keeps getting more expensive: <b>renewals now carry 3&ndash;10% price uplifts year over year</b>, and vendors are adding their own AI features as new, separately billed upsells on top.</p>
    </div>
  </div>

  <div class="card">
    <div class="card-header">
      <h3>The Enterprise Application Value Inversion</h3>
      <p>Click a layer on either side to see where it's headed, then read the executive question below.</p>
    </div>
    <div class="section-grid">
      <div>
        <div class="stack-wrap">
          <div>
            <div class="stack-col-label"><span>Then: value concentrated in</span><span class="dir-down">&#9660; declining</span></div>
            <div class="stack" id="stackThen"></div>
          </div>
          <div class="stack-mid-label">the inversion</div>
          <div>
            <div class="stack-col-label"><span>Now: value concentrated in</span><span class="dir-up">&#9650; rising</span></div>
            <div class="stack" id="stackNow"></div>
          </div>
        </div>
      </div>
      <aside class="detail-panel" id="stackDetail" tabindex="0"></aside>
    </div>
    <div class="footer-strip">If AI can access the data without opening the application, where does the application's value now reside?</div>
  </div>

  <div class="card">
    <div class="card-header">
      <h3>Why costs keep climbing</h3>
      <p>Premium application vendors are not hyperscalers; they rent compute from them. Click each row for the structural comparison.</p>
    </div>
    <div class="section-grid">
      <div id="compareRows"></div>
      <aside class="detail-panel" id="compareDetail" tabindex="0"></aside>
    </div>
  </div>

  <div class="card">
    <div class="card-header">
      <h3>The compounding cost</h3>
      <p>A single vendor's &ldquo;AI upgrade&rdquo; can cost more than the platform itself.</p>
    </div>
    <div class="stat-grid">
      <div class="stat-card"><div class="stat-num">24&ndash;26%</div><p>Typical license cost increase to bolt on a separate data + AI layer</p></div>
      <div class="stat-card" style="border-top-color:var(--cp-purple)"><div class="stat-num" style="color:var(--cp-purple)">$2&ndash;3M</div><p>Typical cost to implement that standalone data layer</p></div>
      <div class="stat-card" style="border-top-color:var(--cp-teal)"><div class="stat-num" style="color:var(--cp-teal)">$1M+ /yr</div><p>Typical ongoing cost just to maintain it, every year</p></div>
    </div>
    <div class="footer-strip">This is a per-vendor tax. Providence likely pays it more than once.</div>
  </div>
</section>
"""

PANEL_STATUS_QUO = f"""
<section id="status-quo" class="panel" aria-labelledby="status-quo-title">
  <div class="panel-head">
    <h2 id="status-quo-title">This is the status quo, and it doesn't work.</h2>
    <p>AI is only as smart as the context it can reach. Most enterprises hand their Copilot one connector into a fragmented stack.</p>
  </div>

  <div class="card">
    <div class="card-header"><h3>The blind spot</h3><p>Most Copilots answer questions. Few actually know your business.</p></div>
    <img class="deck-screenshot" src="{IMG_BLIND_SPOT}" alt="The familiar picture: one connected system, everything else stays a silo Copilot can't see. The fuller picture: every system connected, nothing hidden from view." />
    <div class="footer-strip">None of this requires ripping anything out. AI can extend what you already run today (Epic, Salesforce, Workday, and more) on one shared foundation, with no separate layer to license or maintain.</div>
  </div>

  <div class="card">
    <div class="card-header"><h3>Fragmented context is why most enterprise AI never pays off</h3><p>The data proves it. Every handoff in the care journey feels it.</p></div>
    <div class="stat-grid">
      <div class="stat-card" style="border-top-color:var(--cp-danger)"><div class="stat-num" style="color:var(--cp-danger)">95%</div><p>of generative AI pilots deliver no measurable return; most can't see enough of the business to ground their answers (MIT NANDA, 2025)</p></div>
      <div class="stat-card"><div class="stat-num">$20B</div><p>in annual U.S. healthcare savings still on the table from replacing manual, disconnected admin workflows (CAQH Index, 2025)</p></div>
      <div class="stat-card" style="border-top-color:var(--cp-teal)"><div class="stat-num" style="color:var(--cp-teal)">13 hrs/wk</div><p>average clinician and staff time lost to prior authorization and manual documentation (AMA, 2024)</p></div>
    </div>
    <div class="handoff-row">
      <div class="handoff-card"><b>Pre-diagnosis</b>Patient calls with symptoms; contact center can't see their care history <span class="x">&#10005;</span></div>
      <div class="handoff-card"><b>Diagnosis</b>Clinician re-enters data the EHR and referral system don't share <span class="x">&#10005;</span></div>
      <div class="handoff-card"><b>Treatment</b>Care team coordinates across systems that don't talk to each other <span class="x">&#10005;</span></div>
      <div class="handoff-card"><b>Survivorship</b>Follow-up care depends on someone remembering to make the call <span class="x">&#10005;</span></div>
    </div>
    <div class="footer-strip">The problem was never the model. It's what the model can't see, and it costs trust at the exact moments Providence's mission of human connection depends on it most.</div>
    <p class="source-note">Industry-wide research findings, not Providence-specific data. Care-journey pattern is illustrative, based on common cross-system handoff points in large health systems.</p>
  </div>
</section>
"""

PANEL_BIG_QUESTION = f"""
<section id="big-question" class="panel" aria-labelledby="big-question-title">
  <h2 id="big-question-title" class="sr-only">The Big Question</h2>
  <div class="big-question-panel">
    <img src="{IMG_BIG_QUESTION}" alt="The burner under the pot is already on. The enterprise application value inversion is already under way. In five years, will you look back and wish you'd taken a different strategy?" />
  </div>
</section>
"""

PANEL_OPPORTUNITY = f"""
<section id="opportunity" class="panel" aria-labelledby="opportunity-title">
  <div class="panel-head">
    <h2 id="opportunity-title">What this is worth to Providence.</h2>
    <p>Grounded in Providence's own estimated software spend, not generic industry benchmarks.</p>
  </div>

  <div class="card">
    <div class="card-header"><h3>The AI Business Applications TAM at Providence</h3><p>~{fmt_big_m(TAM_TOTAL)} of AI-BPS-addressable third-party software spend across Providence's application portfolio. Click a category for representative vendors.</p></div>
    <div class="section-grid">
      <div class="pick-grid" id="tamGrid"></div>
      <aside class="detail-panel" id="tamDetail" tabindex="0"></aside>
    </div>
  </div>

  <div class="card">
    <div class="card-header"><h3>The {fmt_big_m(TAM_TOTAL)} opportunity</h3><p>What a 10%-a-year migration to Microsoft saves: annual cost savings vs. today's spend, Year 1&ndash;10.</p></div>
    <div class="chart-wrap">
      <div class="chart-bars" id="savingsChart"></div>
      <div class="chart-legend">
        <span><i style="background:var(--cp-green)"></i>Savings if narrowly licensed (Power Apps only)</span>
        <span><i style="background:var(--cp-accent)"></i>Savings if broadly licensed (Power Apps + Dynamics 365)</span>
      </div>
    </div>
    <div class="stat-grid" style="margin-top:18px">
      <div class="stat-card" style="border-top-color:var(--cp-green)"><div class="stat-num" style="color:var(--cp-green)">{fmt_big_m(FLOOR_SAVINGS_POTENTIAL)}</div><p>Save up to this much every year by Year 10, once fully migrated (as low as {fmt_big_m(CEILING_SAVINGS_POTENTIAL)}/yr if broadly licensed)</p></div>
      <div class="stat-card" style="border-top-color:var(--cp-green)"><div class="stat-num" style="color:var(--cp-green)">{fmt_big_m(TENYR_FLOOR_CUMULATIVE)}</div><p>Up to this much in cumulative 10-year savings (as low as {fmt_big_m(TENYR_CEILING_CUMULATIVE)} if broadly licensed)</p></div>
    </div>
    <p class="source-note">Figures are NET savings: third-party spend avoided minus the Microsoft license cost incurred, modeled against Providence's full 125,000-caregiver-and-staff headcount at $1.85/user/month (Power Apps Premium) and $33.44/user/month (Dynamics 365). Illustrative, not a quote. Actual pace, mix, and licensed seat count will vary by workload.</p>
  </div>

  <div class="card">
    <div class="card-header"><h3>Apps are the real estate: lower cost to own, more capacity for care</h3><p>What this frees up for Providence.</p></div>
    <div class="stat-grid">
      <div class="stat-card" style="border-top-color:var(--cp-green)"><div class="stat-num" style="color:var(--cp-green)">{fmt_big_m(TENYR_CEILING_CUMULATIVE)}&ndash;{fmt_big_m(TENYR_FLOOR_CUMULATIVE)}</div><p>saved over 10 years by migrating to Microsoft: capacity to reinvest in care, not licensing</p></div>
      <div class="stat-card" style="border-top-color:var(--cp-purple)"><div class="stat-num" style="color:var(--cp-purple)">3 &rarr; 1</div><p>application categories (CRM, ERP, App Dev) consolidated onto one platform: less integration risk, faster to deploy new capabilities</p></div>
      <div class="stat-card" style="border-top-color:var(--cp-teal)"><div class="stat-num" style="color:var(--cp-teal)">Time back</div><p>AI agents take on high-volume clinical and admin tasks, freeing staff to focus on patients, not systems</p></div>
    </div>
    <div class="footer-strip">Every dollar not spent on licensing markups and vendor sprawl is a dollar available for clinical capacity, patient experience, and Providence's mission of human connection.</div>
  </div>
</section>
"""

PANEL_NEW_APPROACH = f"""
<section id="new-approach" class="panel" aria-labelledby="new-approach-title">
  <div class="panel-head">
    <h2 id="new-approach-title">The blue ocean strategy: one platform, five connected IQs.</h2>
    <p>Not a better version of the same fragmented approach: a fundamentally different starting point. Real intelligence needs these things connected, not five vendors guessing alone.</p>
  </div>

  <div class="card">
    <div class="card-header"><h3>Your enterprise data, unlocked</h3><p>Click each IQ to see what it connects; the detail sits below, centralized, no matter which you pick.</p></div>
    <div class="pick-grid-wide" id="fiveThingsGrid"></div>
    <aside class="detail-panel detail-panel-below" id="fiveThingsDetail" tabindex="0"></aside>
    <div class="footer-strip">Connected together, this is where real intelligence lives; most vendors only ever give you one.</div>
  </div>

  <div class="card">
    <div class="card-header"><h3>One platform, built to run all five, together</h3><p>The same Power Platform and business application investment, now the foundation your AI runs on too.</p></div>
    <img class="deck-screenshot" src="{IMG_IQ_PLATFORM}" alt="Microsoft IQ platform architecture: Copilot and Workforce layers on top; Work IQ, Customer IQ, Fabric IQ, Foundry IQ, and Web IQ as the intelligence layer; Power Platform and business applications by line of business; Microsoft Fabric / OneLake reaching across every cloud; governance and trust layers underneath; Azure as the foundation." />
    <p class="source-note">Architecture simplified for this canvas. Work IQ, Customer IQ, Fabric IQ, Foundry IQ, and Web IQ are Microsoft's current published (and illustrative, in Customer IQ's case) platform terminology; the business-application modules shown are illustrative of Dynamics 365's line-of-business coverage. Validate exact packaging and availability with the account team.</p>
  </div>

  <div class="card">
    <div class="card-header"><h3>Proof it works</h3><p>This isn't theoretical. Grounded AI is already changing how care gets delivered; these are industry-wide patterns, not Providence-specific results yet.</p></div>
    <div class="stat-grid" id="proofGrid"></div>
  </div>

  <div class="card">
    <div class="card-header"><h3>The compounding return</h3><p>Total cost of ownership goes down. Value realization goes up.</p></div>
    <div class="compound-chart-wrap">
      <svg viewBox="0 0 800 260" preserveAspectRatio="xMidYMid meet" role="img" aria-label="Chart showing total cost of ownership declining from {TCO_START_LABEL} per year to {TCO_END_LABEL} over 10 years, while illustrative value realization rises over the same period, crossing around year six.">
        <polyline points="{TCO_SVG_POINTS}" fill="none" stroke="var(--cp-green)" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round" />
        <polyline points="{VALUE_SVG_POINTS}" fill="none" stroke="var(--cp-purple)" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round" />
        <circle cx="50" cy="{TCO_SVG_POINTS.split()[0].split(',')[1]}" r="5" fill="var(--cp-green)" />
        <circle cx="750" cy="{TCO_SVG_POINTS.split()[-1].split(',')[1]}" r="5" fill="var(--cp-green)" />
        <circle cx="50" cy="{VALUE_SVG_POINTS.split()[0].split(',')[1]}" r="5" fill="var(--cp-purple)" />
        <circle cx="750" cy="{VALUE_SVG_POINTS.split()[-1].split(',')[1]}" r="5" fill="var(--cp-purple)" />
        <text x="50" y="{float(TCO_SVG_POINTS.split()[0].split(',')[1]) - 10}" font-size="15" font-weight="700" fill="var(--cp-green)">{TCO_START_LABEL}/yr</text>
        <text x="620" y="{float(TCO_SVG_POINTS.split()[-1].split(',')[1]) + 20}" font-size="15" font-weight="700" fill="var(--cp-green)">{TCO_END_LABEL}</text>
        <text x="50" y="234" font-size="12" font-style="italic" fill="var(--cp-text-soft)">Yr 1</text>
        <text x="710" y="234" font-size="12" font-style="italic" fill="var(--cp-text-soft)">Yr 10</text>
      </svg>
      <div class="compound-chart-legend">
        <span><i style="background:var(--cp-green)"></i>Total cost of ownership (narrow-licensed floor scenario)</span>
        <span><i style="background:var(--cp-purple)"></i>Value realization &amp; business process automation (illustrative, unitless)</span>
      </div>
    </div>
    <img class="deck-screenshot" style="margin-top:14px" src="{IMG_ENTERPRISE_DATA}" alt="Your enterprise data, unlocked: Your IQ sits on top of Microsoft IQ, which spans Work IQ, Customer IQ, Fabric IQ, Foundry IQ, and Web IQ. The same investment that lowers your cost also compounds your intelligence, and gets more valuable every year you stay on it." />
    <div class="footer-strip">Only Microsoft delivers both from one investment: Dynamics 365 and Power Platform lower the cost; Work IQ, Customer IQ, Fabric IQ, Foundry IQ, and Web IQ compound the value.</div>
  </div>
</section>
"""

PANEL_PLAN = f"""
<section id="plan" class="panel" aria-labelledby="plan-title">
  <div class="panel-head">
    <h2 id="plan-title">Scrutiny is the easy part. Here's the harder question.</h2>
    <p>Recognizing that legacy application contracts deserve a harder look is prudent, but it's not a strategy.</p>
  </div>
  <div class="card">
    <div class="reframe-grid">
      <div class="reframe-box scrutiny">
        <span class="reframe-tag">Part 1 &middot; What scrutiny surfaces</span>
        <div class="reframe-num">{fmt_big_m(TAM_TOTAL)}</div>
        <p style="margin-top:8px;font-size:0.9rem">in legacy third-party spend worth a closer look at your next renewal, with a modeled range of up to {fmt_big_m(TENYR_FLOOR_CUMULATIVE)} in potential savings over 10 years if you migrate to Microsoft.</p>
      </div>
      <div class="reframe-arrow">&rarr;</div>
      <div class="reframe-box matters">
        <span class="reframe-tag">Part 2 &middot; The question that matters more</span>
        <p style="font-size:1.05rem;font-weight:700;margin-top:4px">What if the same platform that could lower your cost also makes every agent smarter, because it finally sees your whole business at once?</p>
      </div>
    </div>
    <div class="footer-strip">The subsequent conversation isn't about a bigger license discount. It's about what your data can finally do once it's not scattered across dozens of vendors.</div>
  </div>

  <div class="card">
    <div class="card-header"><h3>What it takes to get there</h3><p>One decision, two compounding returns, but only if you actually make the shift.</p></div>
    <div class="get-there-list">
      <div class="get-there-step">
        <span class="get-there-badge" style="background:#64748B">1</span>
        <div><h5>Avoid the sunk-cost fallacy</h5><p>Stop protecting legacy apps because of what's already been spent on them.</p></div>
      </div>
      <div class="get-there-step">
        <span class="get-there-badge" style="background:#00338E">2</span>
        <div><h5>Scrutinize with intention</h5><p>Identify where total cost of ownership can actually come down.</p></div>
      </div>
      <div class="get-there-step">
        <span class="get-there-badge" style="background:#39892F">3</span>
        <div><h5>Reorganize the data estate</h5><p>Knock down application and data silos to extract more value.</p></div>
      </div>
    </div>
  </div>

  <div class="close-card">
    <div class="close-logos">
      <img src="{PROVIDENCE_WHITE}" alt="Providence" style="height:30px" />
    </div>
    <h3><span class="muted">The strategic question is no longer <em>which application to buy</em>.</span><br/>It's which platform lowers your cost <span class="hi">and</span> compounds your intelligence, from the same investment.</h3>
    <div class="next-step-box">
      <div class="tag">NEXT STEP</div>
      <p>A complimentary Application Portfolio TCO + AI Readiness Assessment, sized to Providence's actual application estate and data landscape.</p>
    </div>
  </div>

  <p class="source-note" style="margin-top:8px">All cost figures on this canvas (license increase %, implementation cost, maintenance cost, license savings %, TAM figures) are illustrative ranges based on Microsoft field experience and Providence-provided data as of 2026. They are not audited, vendor-published, or contractually guaranteed. The &ldquo;3&ndash;10% YoY SaaS renewal uplift&rdquo; figure and the term &ldquo;SaaSpocalypse&rdquo; reflect widely discussed industry commentary, not audited vendor data. Recommended next step: a joint TCO workshop using actual vendor invoices and contract terms to replace these illustrative figures with validated numbers.</p>
</section>
"""

BODY = HERO + TABS_NAV + PANEL_PROBLEM + PANEL_STATUS_QUO + PANEL_BIG_QUESTION + PANEL_OPPORTUNITY + PANEL_NEW_APPROACH + PANEL_PLAN
print("BODY length:", len(BODY))

# =============================================================== JS ========
JS = """
// -------- data --------------------------------------------------------
const stackThenItems = [
  { label: "User Interfaces", weight: 34, color: "#00338E", detail: "Screens and forms were the product for two decades. That's exactly the layer AI agents route around first." },
  { label: "Workflows", weight: 24, color: "#1464C8", detail: "Step-by-step click paths built for humans lose relevance once an agent can execute the underlying task directly." },
  { label: "Reporting", weight: 18, color: "#2E5A96", detail: "Dashboards summarize data for a human to read. An agent doesn't need a summary; it needs the data itself." },
  { label: "Custom Screens", weight: 13, color: "#16344f", detail: "Bespoke UI built for one team's workflow is the most expensive kind of value to maintain, and the first to go stale." },
  { label: "Embedded Analytics", weight: 11, color: "#0d2038", detail: "Point-in-time analytics modules are being absorbed into always-on AI orchestration across the whole estate." },
];
const stackNowItems = [
  { label: "Data", weight: 34, color: "#071835", textColor: "#fff", detail: "Whoever controls clean, unified data controls what every future AI agent can actually do. This is the new center of gravity." },
  { label: "Context & Knowledge", weight: 24, color: "#00338E", detail: "Policies, history, and institutional memory are what make an answer trustworthy instead of merely plausible." },
  { label: "Compute", weight: 13, color: "#39892F", textColor: "#fff", detail: "Every AI action costs compute. Whoever owns the compute layer captures the margin, or pays someone else's markup for it." },
  { label: "AI Orchestration", weight: 11, color: "#5B3A8E", textColor: "#fff", detail: "Coordinating agents across systems, not just answering one question at a time, is where the next wave of value concentrates." },
];

const compareRows = [
  {
    label: "PREMIUM SaaS (e.g., Salesforce, Workday, Epic)",
    title: "Application License \\u2192 Rented Compute \\u2192 Your Bill",
    sub: "AWS \\u00b7 GCP \\u00b7 other hyperscalers, plus markup",
    bars: [{ pct: 65, color: "#9AA5B1", label: "license" }, { pct: 35, color: "#A14545", label: "+ markup" }],
    detail: "Salesforce, Workday, Epic, and similar premium application vendors run their platforms on third-party cloud infrastructure rather than operating their own hyperscale cloud. That rented-compute cost gets passed through to you, on top of the license.",
  },
  {
    label: "MICROSOFT",
    title: "Application + Cloud: one company",
    sub: "Hyperscaler and business applications vendor, together",
    bars: [{ pct: 100, color: "#008575", label: "license only, no markup slice" }],
    detail: "Microsoft is the leading vendor that is both a hyperscaler and a business applications company, eliminating the compute pass-through markup embedded in premium SaaS contracts like these. (Oracle also spans both, but Microsoft is the most complete version of that model at enterprise scale.)",
  },
];

const fiveThings = [
  { title: "How your people work", meta: "Work IQ", color: "#00338E", detail: "Roles, workflows, files, meetings, and the day-to-day rhythm of the organization. This is the context most Copilots already have some access to." },
  { title: "How you understand your patients", meta: "Customer IQ", color: "#B8792A", detail: "A unified patient profile with engagement signals: care history, preferences, and outreach, connected instead of scattered across a contact center, a CRM, and a marketing platform." },
  { title: "How your business runs", meta: "Fabric IQ", color: "#39892F", detail: "Data, semantics, rules, and the live signals the business runs on: claims, encounters, supply chain, finance. Most AI never reaches this layer at all." },
  { title: "What your org knows", meta: "Foundry IQ", color: "#5B3A8E", detail: "Policies, documents, research, and the knowledge locked in institutional memory; usually scattered across SharePoint, wikis, and people's heads." },
  { title: "What's happening beyond your walls", meta: "Web IQ", color: "#1464C8", detail: "Public and licensed signal from outside the four walls of the organization: market, regulatory, and competitive context most internal AI never considers." },
];

const tamItems = [
  { title: "Low-Code App Dev & Automation", meta: "$23.7M addressable", color: "#5B3A8E",
    vendors: ["ServiceNow", "UiPath", "Automation Anywhere", "Blue Prism"],
    detail: "Automation and low-code tooling spend that could consolidate onto one shared, low-code application platform instead of several standalone automation vendors." },
  { title: "Member & Engagement Platforms", meta: "$86.0M addressable", color: "#00338E",
    vendors: ["Salesforce", "Genesys", "Marketo", "NICE", "Avaya", "Cisco"],
    detail: "Core CRM, marketing, and sales software spend: the largest single category, spanning patient engagement, contact center, and marketing platforms that each carry their own rented-compute markup and isolated data layer." },
  { title: "Finance & Operations", meta: "$142.4M addressable", color: "#39892F",
    vendors: ["SAP", "Oracle", "Workday", "Kronos", "Infor", "Concur"],
    detail: "ERP, HR, and procurement systems spend: the single largest addressable category, and prime candidates for consolidation onto one shared data + AI foundation. Excludes Manufacturing PLM spend, which has no direct Dynamics 365 replacement." },
];

const proofItems = [
  { title: "Ambient clinical documentation", detail: "Reduces time clinicians spend writing notes after each visit, giving time back for patients instead of paperwork." },
  { title: "Automated prior authorization", detail: "Cuts the manual, cross-system work behind hours lost to prior authorization every week, by grounding the request in data that already exists." },
  { title: "Unified patient & contact center context", detail: "Agents and AI see the same case history at once, instead of asking patients to repeat themselves across departments." },
  { title: "Automated claims & revenue cycle", detail: "Reduces manual rework on denials and underpayments by grounding claims and coding in data that already exists across systems." },
  { title: "Coordinated care management", detail: "Care teams and AI agents share the same care plan in real time, closing gaps at every handoff instead of after the fact." },
  { title: "Supply chain visibility", detail: "Real-time inventory visibility helps prevent costly stockouts and waste across facilities and departments." },
];

// -------- generic pick-group builder ------------------------------------
function buildPickGroup(containerId, detailId, items, kicker, opts) {
  opts = opts || {};
  const container = document.getElementById(containerId);
  const detail = document.getElementById(detailId);
  if (!container || !detail) return;
  let active = 0;
  function render() {
    container.innerHTML = "";
    items.forEach((item, i) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "pick-btn" + (i === active ? " active" : "");
      const idxHtml = opts.numbered
        ? `<span class="pick-index" style="background:${item.color}">${i + 1}</span>`
        : `<span class="pick-index" style="background:${item.color}">&#9679;</span>`;
      btn.innerHTML = idxHtml + `<span class="pick-title">${item.title}</span>` + (item.meta ? `<span class="pick-meta">${item.meta}</span>` : "");
      btn.addEventListener("click", () => select(i));
      container.appendChild(btn);
    });
  }
  function select(i) {
    active = i;
    render();
    const item = items[i];
    let extra = "";
    if (item.vendors) {
      extra = `<div class="tag-row">` + item.vendors.map(v => `<span class="detail-tag">${v}</span>`).join("") + `</div>`;
    }
    detail.innerHTML = `<p class="kicker">${kicker}</p><h4>${item.title}</h4><p>${item.detail}</p>${extra}`;
  }
  render();
  select(0);
}

// -------- Then/Now value stack ------------------------------------------
function buildStack(containerId, items) {
  const el = document.getElementById(containerId);
  if (!el) return;
  const total = items.reduce((s, it) => s + it.weight, 0);
  items.forEach((item, i) => {
    const seg = document.createElement("div");
    seg.className = "stack-seg";
    seg.style.background = item.color;
    seg.style.color = item.textColor || "#fff";
    seg.style.flex = (item.weight / total).toString();
    seg.textContent = item.label;
    seg.dataset.index = i;
    el.appendChild(seg);
  });
}

function wireStackClicks() {
  const detail = document.getElementById("stackDetail");
  if (!detail) return;
  document.querySelectorAll("#stackThen .stack-seg, #stackNow .stack-seg").forEach((seg) => {
    seg.addEventListener("click", () => {
      document.querySelectorAll("#stackThen .stack-seg, #stackNow .stack-seg").forEach(s => s.classList.remove("active"));
      seg.classList.add("active");
      const isThen = seg.closest("#stackThen") !== null;
      const idx = Number(seg.dataset.index);
      const item = (isThen ? stackThenItems : stackNowItems)[idx];
      detail.innerHTML = `<p class="kicker">${isThen ? "Then: declining" : "Now: rising"}</p><h4>${item.label}</h4><p>${item.detail}</p>`;
    });
  });
}

// -------- hyperscaler compare rows --------------------------------------
function buildCompareRows() {
  const container = document.getElementById("compareRows");
  const detail = document.getElementById("compareDetail");
  if (!container || !detail) return;
  let active = 0;
  function render() {
    container.innerHTML = "";
    compareRows.forEach((row, i) => {
      const div = document.createElement("div");
      div.className = "compare-row" + (i === active ? " active" : "");
      const barHtml = row.bars.map(b => `<span style="width:${b.pct}%;background:${b.color}">${b.label}</span>`).join("");
      div.innerHTML = `<span class="row-label">${row.label}</span><span class="row-title">${row.title}</span><span class="row-sub">${row.sub}</span><div class="bill-bar">${barHtml}</div>`;
      div.addEventListener("click", () => select(i));
      container.appendChild(div);
    });
  }
  function select(i) {
    active = i;
    render();
    const row = compareRows[i];
    detail.innerHTML = `<p class="kicker">Structural comparison</p><h4>${row.label}</h4><p>${row.detail}</p>`;
  }
  render();
  select(0);
}

// -------- savings bar chart ----------------------------------------------
function buildSavingsChart() {
  const el = document.getElementById("savingsChart");
  if (!el) return;
  const floor = FLOOR_ANNUAL;
  const ceiling = CEILING_ANNUAL;
  const maxVal = Math.max(...floor, ...ceiling);
  floor.forEach((floorVal, i) => {
    const ceilingVal = ceiling[i];
    const year = document.createElement("div");
    year.className = "chart-year";
    const pair = document.createElement("div");
    pair.className = "chart-bar-pair";
    const narrowBar = document.createElement("div");
    narrowBar.className = "chart-bar narrow";
    narrowBar.style.height = (floorVal / maxVal * 100) + "%";
    narrowBar.innerHTML = `<span class="bar-label">$${Math.round(floorVal)}M</span>`;
    const broadBar = document.createElement("div");
    broadBar.className = "chart-bar broad";
    broadBar.style.height = (ceilingVal / maxVal * 100) + "%";
    broadBar.innerHTML = `<span class="bar-label">$${Math.round(ceilingVal)}M</span>`;
    pair.appendChild(narrowBar);
    pair.appendChild(broadBar);
    year.appendChild(pair);
    const label = document.createElement("div");
    label.className = "chart-year-label";
    label.textContent = "Yr " + (i + 1);
    year.appendChild(label);
    el.appendChild(year);
  });
}

// -------- proof grid (static stat cards) ---------------------------------
function buildProofGrid() {
  const el = document.getElementById("proofGrid");
  if (!el) return;
  const palette = ["#00338E", "#39892F", "#5B3A8E", "#C9862E", "#1464C8", "#008575"];
  proofItems.forEach((item, i) => {
    const card = document.createElement("div");
    card.className = "stat-card";
    card.style.borderTopColor = palette[i % palette.length];
    card.innerHTML = `<h5 style="margin:0 0 8px;font-size:0.98rem">${item.title}</h5><p>${item.detail}</p>`;
    el.appendChild(card);
  });
}

// -------- tooltips ---------------------------------------------------------
function wireTooltips() {
  const tip = document.createElement("div");
  tip.className = "tooltip";
  document.body.appendChild(tip);
  function show(btn) {
    tip.innerHTML = `<div class="tip-title">${btn.dataset.tipTitle}</div>${btn.dataset.tip}`;
    const rect = btn.getBoundingClientRect();
    tip.style.left = rect.left + "px";
    tip.style.top = (rect.bottom + 8) + "px";
    tip.classList.add("visible");
  }
  function hide() { tip.classList.remove("visible"); }
  document.querySelectorAll("[data-tip]").forEach((btn) => {
    btn.addEventListener("mouseenter", () => show(btn));
    btn.addEventListener("mouseleave", hide);
    btn.addEventListener("focus", () => show(btn));
    btn.addEventListener("blur", hide);
  });
}

// -------- tab wiring --------------------------------------------------------
function wireTabs() {
  const buttons = document.querySelectorAll(".tab-button");
  const panels = document.querySelectorAll(".panel");
  buttons.forEach((btn) => {
    btn.addEventListener("click", () => {
      buttons.forEach(b => b.classList.remove("active"));
      panels.forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(btn.dataset.tab).classList.add("active");
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  });
}

// -------- theme-aware hero logo swap --------------------------------------
function wireLogoTheme() {
  const dark = document.documentElement.getAttribute("data-theme") === "dark";
  if (!dark) return;
  const providence = document.getElementById("logo-providence");
  if (providence) providence.src = window.__PROVIDENCE_LOGO_WHITE__;
}

// -------- init --------------------------------------------------------
wireLogoTheme();
buildStack("stackThen", stackThenItems);
buildStack("stackNow", stackNowItems);
wireStackClicks();
buildCompareRows();
buildPickGroup("fiveThingsGrid", "fiveThingsDetail", fiveThings, "Connected foundation");
buildPickGroup("tamGrid", "tamDetail", tamItems, "TAM category", { numbered: false });
buildSavingsChart();
buildProofGrid();
wireTabs();
wireTooltips();
"""
print("JS length:", len(JS))

# Inject the computed year-by-year savings arrays as JS constants ahead of
# the functions that consume them (kept out of the big triple-quoted JS
# string above so the numbers stay driven by the Python model).
_FLOOR_ANNUAL_JS = "[" + ",".join(f"{v:.2f}" for v in TENYR_FLOOR_ANNUAL) + "]"
_CEILING_ANNUAL_JS = "[" + ",".join(f"{v:.2f}" for v in TENYR_CEILING_ANNUAL) + "]"
JS = f"const FLOOR_ANNUAL = {_FLOOR_ANNUAL_JS};\nconst CEILING_ANNUAL = {_CEILING_ANNUAL_JS};\n" + JS

# ========================================================= assemble =======
HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Providence &middot; The Enterprise Application Value Inversion</title>
<style>{CSS}</style>
<script>
window.__PROVIDENCE_LOGO_WHITE__ = "{PROVIDENCE_WHITE}";
</script>
</head>
<body>
<div class="app">
{BODY}
</div>
<footer class="hub-footer">&copy; 2026 Microsoft &middot; Prepared for Providence &middot; Not for external redistribution</footer>
<script>
{JS}
</script>
</body>
</html>
"""

OUT_PATH = os.path.join(RAW_DIR, "decision-canvas.html")
with open(OUT_PATH, "w", encoding="utf-8") as f:
    f.write(HTML)
print(f"Wrote {OUT_PATH} ({len(HTML)} chars)")
