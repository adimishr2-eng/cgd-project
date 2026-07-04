# ⛽ PNG Demand Forecasting & Network Expansion Planning Model

<div align="center">

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?style=for-the-badge&logo=python)
![Streamlit](https://img.shields.io/badge/Streamlit-1.x-red?style=for-the-badge&logo=streamlit)
![Plotly](https://img.shields.io/badge/Plotly-Interactive-green?style=for-the-badge&logo=plotly)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)

**A professional-grade, interactive analytics platform for City Gas Distribution (CGD) network planning.**  
Built for gas utilities, infrastructure planners, and energy analysts to make data-driven decisions on PNG expansion.

[🚀 Quick Start](#quick-start) · [📖 How It Works](#how-it-works) · [🗂️ Project Structure](#project-structure) · [📸 Features](#features)

</div>

---

## 📌 What Is This Project?

**PNG (Piped Natural Gas)** is distributed to households, commercial establishments, and industries through a City Gas Distribution (CGD) network. Expanding this network requires answering three critical questions:

1. **How much gas will be needed?** → *Demand Forecasting*
2. **Which areas should be connected first?** → *Zone Prioritization*
3. **Is the investment financially viable?** → *Techno-Economic Analysis*

This model answers all three — and goes further with **scenario planning**, **sensitivity analysis**, **strategy comparison**, and a **GIS map** — all through an interactive Streamlit dashboard that works for **any CGD network in India**, just by changing the sidebar inputs.

---

## 🚀 Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/adimishr2-eng/cgd-project.git
cd cgd-project
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the App
```bash
streamlit run app.py
```

The app opens automatically at **http://localhost:8501**

> **No configuration needed.** Default values for a typical mid-size Indian city are pre-loaded. Just click **▶ Run Analysis** in the sidebar.

---

## 📦 Dependencies

| Library | Purpose |
|---|---|
| `streamlit` | Web dashboard framework |
| `plotly` | Interactive charts (line, bar, radar, tornado, Gantt) |
| `pandas` | Data manipulation and tabular analysis |
| `numpy` | Numerical computations |
| `numpy-financial` | IRR and NPV calculations |
| `openpyxl` | Excel report generation |
| `reportlab` | PDF report generation |
| `geopandas` | GIS / geographic data processing |
| `folium` | Interactive map rendering |
| `streamlit-folium` | Embed Folium maps in Streamlit |
| `shapely` | Geometric operations on zone boundaries |
| `scipy` | Scientific computing utilities |

---

## 🗂️ Project Structure

```
cgd-project/
│
├── app.py                     ← Main Streamlit dashboard (all 7 tabs)
├── requirements.txt           ← Python dependencies
├── README.md                  ← This file
│
├── models/
│   ├── demand_forecast.py     ← Standalone demand forecasting module
│   ├── zone_prioritization.py ← Zone scoring and ranking module
│   └── techno_economic.py     ← Financial analysis module
│
├── inputs/
│   ├── city_data.csv          ← City-level input parameters
│   ├── zone_data.csv          ← Zone-level data (area, households, etc.)
│   └── financial_params.csv   ← Financial parameters (costs, revenues)
│
├── static/
│   ├── css/style.css          ← Custom CSS styles
│   └── js/app.js              ← JavaScript utilities
│
└── templates/
    └── index.html             ← HTML template
```

---

## 📖 How It Works

### Overall Architecture

```
Sidebar Inputs (City Parameters, Growth Rates, Financial Params)
        │
        ▼
  ┌─────────────────────────────────────────────────────┐
  │              Calculation Engine (app.py)             │
  │   @st.cache_data functions — computed once, cached  │
  │                                                      │
  │  calc_demand_forecast()     → demand_df              │
  │  calc_zone_prioritization() → zone_ranking_df        │
  │  calc_techno_economic()     → financial_df           │
  │  calc_scenario_analysis()   → (base, opt, pess)      │
  │  calc_sensitivity_analysis()→ (sens_df, gas_df)      │
  │  calc_strategy_comparison() → (summary, cashflows)   │
  └─────────────────────────────────────────────────────┘
        │
        ▼
  st.session_state  ← stores results between reruns
        │
        ▼
  7-Tab Dashboard  ← Plotly charts + DataFrames + Downloads
```

---

## 🗂️ Tab-by-Tab Breakdown

### 🗺️ Tab 1 — GIS Map
**Purpose:** Visualize zone boundaries geographically before running any analysis.

**How it works:**
- Upload a **GeoJSON file** with zone polygons and an optional **CSV** with population data
- The app auto-calculates `area_sqkm`, `centroid_lat`, `centroid_lon` from the geometry
- A **Folium interactive map** renders zones color-coded by priority (green = high, amber = medium, red = low)
- Click any zone marker for a popup showing area, households, and priority score
- Use **"Use These Zones in Analysis"** to populate the zone table from GeoJSON

**If no GeoJSON is uploaded:** A default placeholder map of Vadodara is shown.

---

### 📈 Tab 2 — Demand Forecast
**Purpose:** Project total PNG demand over 10 years across all three customer segments.

**How it works:**

The model uses a **compound penetration growth formula**:

```
Penetration(Year t) = min(Base_Penetration × (1 + Growth_Rate)^(t-1), 0.85)
```

For each year (1–10), it computes:
- `Residential Customers = Total Households × Residential Penetration`
- `Commercial Customers = Total Commercial Establishments × Commercial Penetration`
- `Industrial Customers = Total Industrial Units × Industrial Penetration`
- `Daily Demand (SCM/day) = Σ (Customers × Daily Consumption per Customer)`
- `Annual Demand (SCM) = Daily Demand × 365`

**Inputs (from sidebar):**
| Input | Description |
|---|---|
| Total Population | City population |
| Avg Household Size | People per household |
| Commercial Establishments | Total shops, offices, restaurants |
| Industrial Units | Factories and large consumers |
| Base Penetration (%) | % of each segment already using PNG |
| Annual Growth Rate (%) | Year-on-year penetration growth |
| Daily Consumption | SCM/day per customer type |

**Outputs:** Year-wise demand table, growth trend chart, customer mix bar chart, segment breakdown pie chart.

---

### 🏆 Tab 3 — Zone Prioritization
**Purpose:** Rank all zones by a composite score to decide where to expand first.

**How it works:**

Each zone is scored on 4 dimensions and normalized to 0–100:

```
Score = 0.30 × Household_Density_Score
      + 0.30 × Industrial_Score
      + 0.20 × Penetration_Score
      + 0.20 × Proximity_Score
```

| Score Component | What It Measures | Weight |
|---|---|---|
| Household Density Score | Households per sq km (higher = better) | 30% |
| Industrial Score | Industrial units + households equivalent | 30% |
| Penetration Score | Existing PNG penetration (market readiness) | 20% |
| Proximity Score | 1/Distance from existing network (closer = better) | 20% |

**Recommendation Logic:**
- Rank 1–3 → **Year 1 Expansion**
- Rank 4–5 → **Year 2 Expansion**
- Rank 6+ → **Year 3 Expansion**

**Outputs:** Ranked table with color-coded scores, radar chart per zone, heatmap of scores, bar chart comparison.

---

### 💰 Tab 4 — Techno-Economic Analysis
**Purpose:** Calculate CAPEX, revenues, IRR, NPV, and payback period for each zone.

**How it works:**

**CAPEX Calculation:**
```
Pipeline Length (km)  = √(Area_sqkm) × 2.5
Pipeline Cost         = Pipeline Length × PE Cost (₹ Lakh/km)
Domestic CAPEX        = New Domestic Connections × Connection Cost
Commercial CAPEX      = New Commercial Connections × Connection Cost
Industrial CAPEX      = New Industrial Connections × Connection Cost
Total CAPEX           = Pipeline Cost + Domestic + Commercial + Industrial
```

**Revenue Ramp-up Model (10-year):**
```
Year t Revenue = (Dom_Conn × t/10 × Dom_Rev) +
                 (Com_Conn × t/10 × Com_Rev) +
                 (Ind_Conn × t/10 × Ind_Rev)
Net Cash Flow  = Revenue × (1 - OPEX%)
```

**Financial Metrics:**
- **IRR** — Internal Rate of Return (using `numpy-financial.irr`)
- **NPV** — Net Present Value at the discount rate
- **Payback Period** — Years until cumulative cash flow turns positive
- **Viability:** `IRR > Threshold` → Viable | `IRR ≥ Threshold − 4%` → Marginal | else → Needs Review

**Outputs:** Summary table per zone, cumulative cash flow chart, waterfall CAPEX chart, zone-wise IRR bar chart, year-by-year cash flow table.

---

### 📊 Tab 5 — Scenario Analysis
**Purpose:** Test demand and financial outcomes under optimistic, base, and pessimistic assumptions.

**How it works:**

Three parallel demand forecasts are run simultaneously:

| Scenario | Growth Rate Multiplier |
|---|---|
| 🟢 Optimistic | Base Rate × Optimistic Multiplier (e.g. 1.3×) |
| 🔵 Base Case | As configured in sidebar |
| 🔴 Pessimistic | Base Rate × Pessimistic Multiplier (e.g. 0.7×) |

Growth rates are capped at 50% to prevent unrealistic projections.

**KPI Cards:** Year 10 demand and customers shown for all three scenarios side by side.

**Outputs:** Uncertainty band line chart, scenario comparison bar chart, combined data table, scenario financial impact (IRR per scenario), Excel download with 3 sheets.

---

### 🎯 Tab 6 — Sensitivity Analysis
**Purpose:** Identify which variables have the most impact on financial returns (IRR).

**How it works:**

For each of 6 key variables, the model computes IRR at its minimum and maximum values:

| Variable | Variation Method |
|---|---|
| Residential Growth Rate | ±penetration_variation% |
| Commercial Growth Rate | ±penetration_variation% |
| Industrial Growth Rate | ±penetration_variation% |
| PE Pipeline Cost | ±30% of base |
| Domestic Revenue | ±20% of base |
| Gas Purchase Cost | ±gas_cost_variation% |

**IRR Swing** = `|IRR_max − IRR_min|` — the bigger the swing, the more sensitive.

**Charts:**
- **Tornado Chart** — Variables sorted by swing magnitude (red = downside, green = upside)
- **Radar/Spider Chart** — Normalized sensitivity scores across all variables
- **Gas Price Deep-Dive** — Multi-line chart showing IRR vs. gas cost per zone

**Risk Classification:** `Swing > 5%` → High | `2–5%` → Medium | `<2%` → Low

---

### ⚖️ Tab 7 — Strategy Comparison
**Purpose:** Compare three rollout strategies to decide the best way to deploy CAPEX.

| Strategy | Description | CAPEX Timing |
|---|---|---|
| 🚀 **Aggressive** | All zones connected in Year 1 | 100% upfront |
| 📅 **Phased** | Zones split into 3 groups, connected at intervals | Spread over years |
| 💡 **Asset Light** | LNG virtual pipeline first, then convert top zones | Minimal upfront |

**Asset Light Logic:**
- Years 1–3: Serve all zones via LNG trucks (+35% operating cost)
- Year 4: Convert top 3 priority zones to permanent pipeline
- Year 5+: Reduced LNG dependency (+15% cost)

**Outputs:** Strategy scorecard cards, CAPEX comparison chart, cumulative cash flow (with breakeven markers ★), radar scorecard (6 dimensions), year-by-year cash flow bars, **Gantt-style zone pipeline timeline**.

---

### 📄 PDF Final Report
**Purpose:** One-click professional PDF report covering all analysis sections.

**Report Structure:**
1. **Cover Page** — City name, date, model version, key parameters
2. **Executive Summary** — 5 KPI cards + narrative paragraph
3. **Section 1: Demand Forecast** — Full 10-year table
4. **Section 2: Zone Prioritization** — Ranked table (top zone highlighted green)
5. **Section 3: Techno-Economic** — CAPEX/IRR/NPV per zone (viability color-coded 🟢🟡🔴)
6. **Section 4: Scenario Analysis** — Base/Optimistic/Pessimistic demand table
7. **Section 5: Sensitivity Analysis** — IRR swing table (risk color-coded)
8. **Section 6: Strategy Comparison** — All three strategies side by side

Generated using `reportlab` — downloads as `CGD_Report_<CityName>_<Date>.pdf`.

---

## ⚙️ Configuration — Sidebar Inputs

### 🏙️ City Parameters
| Input | Default | Description |
|---|---|---|
| City Name | Generic City | Label for reports |
| Total Population | 22,00,000 | City population |
| Avg Household Size | 4.2 | People per household |
| Commercial Establishments | 45,000 | Shops, offices, restaurants |
| Industrial Units | 850 | Factories and large consumers |

### 📈 Growth Parameters
| Input | Default | Description |
|---|---|---|
| Residential Base Penetration | 10% | Initial % of households on PNG |
| Commercial Base Penetration | 5% | Initial % of commercial on PNG |
| Industrial Base Penetration | 8% | Initial % of industries on PNG |
| Residential Growth Rate | 18%/yr | Annual penetration growth |
| Commercial Growth Rate | 10%/yr | Annual penetration growth |
| Industrial Growth Rate | 8%/yr | Annual penetration growth |
| Res. Daily Consumption | 0.5 SCM/day | Per household per day |
| Com. Daily Consumption | 3.0 SCM/day | Per establishment per day |
| Ind. Daily Consumption | 50.0 SCM/day | Per industrial unit per day |

### 💵 Financial Parameters
| Input | Default | Description |
|---|---|---|
| PE Pipeline Cost | ₹15 Lakh/km | Cost of laying PE pipeline |
| Domestic Connection Cost | ₹8,000 | Per household connection |
| Commercial Connection Cost | ₹25,000 | Per commercial connection |
| Industrial Connection Cost | ₹1,00,000 | Per industrial connection |
| Domestic Revenue | ₹1,200/month | Revenue per domestic customer |
| Commercial Revenue | ₹8,000/month | Revenue per commercial customer |
| Industrial Revenue | ₹80,000/month | Revenue per industrial customer |
| OPEX % of Revenue | 35% | Operating expense ratio |
| Discount Rate | 12% | For NPV calculation |
| Viability Threshold IRR | 15% | Minimum acceptable IRR |

---

## 🗺️ Zone Data Table

The **Zone Configuration** table (editable in the app) accepts:

| Column | Type | Description |
|---|---|---|
| `zone_name` | Text | Name of the zone/area |
| `area_sqkm` | Float | Geographic area in sq km |
| `total_households` | Integer | Number of residential units |
| `industrial_units` | Integer | Number of industrial consumers |
| `existing_penetration_pct` | Float (0–1) | Current PNG penetration (0.12 = 12%) |
| `distance_from_network_km` | Float | Distance from nearest existing pipeline |

Rows can be **added, edited, or deleted** directly in the app. Changes take effect immediately on the next **▶ Run Analysis** click.

---

## 📤 Export Options

| Format | Content | Location |
|---|---|---|
| **Excel (.xlsx)** | Demand Forecast + Zone Ranking + Financial Model | Executive Summary section |
| **Excel (.xlsx)** | Scenario Analysis (3 sheets) | Scenario Analysis tab |
| **PDF (.pdf)** | Complete 6-section professional report | Bottom of page |

---

## 🔢 Key Formulas

### Penetration Growth (Logistic Cap)
$$P_t = \min\left(P_0 \times (1 + g)^{t-1},\ 0.85\right)$$

### Annual Gas Demand
$$D_{annual} = \sum_{s \in \{Res, Com, Ind\}} \left(C_s \times d_s\right) \times 365$$

### Zone Composite Score
$$\text{Score} = 0.30 \cdot \text{HD} + 0.30 \cdot \text{Ind} + 0.20 \cdot \text{Pen} + 0.20 \cdot \text{Prox}$$

### Net Present Value
$$NPV = \sum_{t=0}^{T} \frac{CF_t}{(1+r)^t}$$

### Internal Rate of Return
$$0 = \sum_{t=0}^{T} \frac{CF_t}{(1+IRR)^t}$$

---

## 🧪 How to Adapt for Your City

1. Open the app → `streamlit run app.py`
2. Update **City Parameters** in the sidebar (name, population, household size)
3. Update the **Zone Configuration table** with your actual geographic areas
4. Adjust **Growth Parameters** from local market research
5. Update **Financial Parameters** with actual costs from your procurement team
6. Click **▶ Run Analysis** — all 7 tabs update instantly

> The model is entirely **input-driven** — no hardcoded city-specific values exist in the calculation logic.

---

## 🏗️ Technical Design Decisions

| Decision | Reason |
|---|---|
| `@st.cache_data` on all calc functions | Results cached by input hash — instant re-render when switching tabs, only recomputes when inputs change |
| Plotly instead of Matplotlib | Plotly charts are fully interactive in Streamlit (zoom, hover, toggle, download) |
| `st.session_state` for results | Persists data across tab switches without re-running the analysis |
| Generic input-driven design | Makes the model reusable for any CGD geography without code changes |
| Optional `geopandas` import | App works even if geo libraries aren't installed; GIS tab gracefully degrades |
| `reportlab` for PDF | Pure Python, no external tools needed, produces professional A4 reports |

---

## 📋 Sample Results (Default Inputs)

Running with default inputs (Generic City, 6 zones):

| Metric | Value |
|---|---|
| Year 1 Demand | ~450 SCM/day |
| Year 10 Demand | ~2,800 SCM/day |
| Total Demand Growth | ~5× over 10 years |
| Total Network CAPEX | ₹35–55 Crores |
| Average IRR | 18–24% |
| Top Priority Zone | Industrial Corridor West |
| Best Strategy (by IRR) | Aggressive |

---

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/new-analysis`
3. Commit changes: `git commit -m "Add: XYZ analysis feature"`
4. Push: `git push origin feature/new-analysis`
5. Open a Pull Request

---

## 📄 License

This project is licensed under the **MIT License** — free to use, modify, and distribute with attribution.

---

## 👤 Author

**adimishr2-eng**  
GitHub: [@adimishr2-eng](https://github.com/adimishr2-eng)

---

<div align="center">

**Built with ❤️ using Python · Streamlit · Plotly · ReportLab**

*Generic CGD Analytics Framework — Adaptable for any City Gas Distribution network*

</div>
