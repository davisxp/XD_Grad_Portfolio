# Methodology — Life‑Sciences Finance Portfolio
**Maintainer:** Xavier Davis • **Updated:** 27 Oct 2025  
Scope: consistent, auditable methods for valuation, modelling, and documentation across all six pitches. This is not investment advice; public sources only.

---

## 1) Core principles
- **Traceability:** every figure in notes links to a cell; PDF and Excel carry the **same version tag** (e.g., v1.2).  
- **Single source of truth:** Excel workbooks are authoritative; CSVs/PDFs are views for web reading.  
- **Reproucibility:** all assumptions are dated and referenced (RNS, Annual Report/10‑K, trial registry).  
- **Conservatism:** base cases avoid heroic assumptions; upside/downside live in scenarios, not the base.
- **Units & currency:** GBP by default (£m), with explicit FX if source is USD/EUR. Units stated on each table.

---

## 2) Valuation frameworks
### 2.1 Clinical‑stage biopharma — **rNPV**
Used when cash flows are contingent on clinical success.
- **Cash flows:** revenue = patients × penetration × price (gross‑to‑net), less COGS and op costs/royalties where relevant.
- **Timing:** launch year + ramp (S‑curve), peak penetration, LOE (loss of exclusivity) via patent/biologic tail.
- **Probability:** phase/indication‑specific PoS (probability of success) applied to expected cash flows.
- **Discounting:** risk‑adjusted **discount rate** (nominal, post‑tax). Baseline: 12% (adjust 9–14% by risk, quality, and rates regime).
- **Cross‑checks:** trading comps (EV/sales, EV/EBITDA where meaningful) and precedent deals.

### 2.2 Commercial‑stage biopharma / medtech — **DCF + Comps**
- **DCF:** 5–10 year explicit with terminal value (g or exit multiple), WACC baseline 8–10% for diversified/medtech, case‑by‑case for pure‑play biotech with marketed assets.
- **Comps:** 5–8 peers; normalise for growth/margins; triangulate to sanity‑check the DCF/rNPV output.

### 2.3 M&A / LBO / Credit
- **Merger:** sources & uses; purchase price allocation (PPA); synergy timing/cost to achieve; debt schedule; accretion/dilution.
- **LBO:** tranches, fees, cash sweep, covenant headroom, MOIC/IRR with entry/exit bridges.
- **LevFin/DCM:** rating‑case vs base‑case metrics (gross/net leverage, ICR, FCF conversion), maturity wall, covenant capacity.

---

## 3) Probability of Success (PoS) — baseline table
Adjust by indication, MoA novelty, trial design, prior data quality. These are **starting points**, not truths.

| Stage | General | Oncology |
|---|---:|---:|
| Preclinical | 5–10% | 3–8% |
| Phase I | 10–20% | 8–15% |
| Phase II | 20–40% | 15–30% |
| Phase III | 50–70% | 45–60% |
| Filed/Regulatory | 80–90% | 75–85% |

**Note:** If robust surrogate endpoints or breakthrough/regenerative pathways apply, tilt upward; if single‑site/small‑N or safety flags, tilt downward.

---

## 4) Discount rate / WACC policy
- **Clinical rNPV:** 12% baseline (nominal, post‑tax); +/- based on asset risk, balance sheet strength, and rate regime.
- **Commercial DCF:** 9–10% typical for established medtech/diversified names; biotech with marketed assets often higher (10–12%).
- **Sensitivity:** always include ±1pp on discount/WACC in the sensitivity set.
Document any deviation explicitly in the model **Assumptions** tab and in the note footnotes.

---

## 5) Revenue drivers & uptake
- **Epidemiology:** prevalence/incidence → addressable population → treatable fraction (contraindications/compliance).  
- **Uptake curve:** logistic/S‑curve; specify time to peak (e.g., 4–6 years).  
- **Pricing:** list price, gross‑to‑net (rebates, access), geography mix if used.  
- **Competition:** show a 2–3 line competitive map and adjust penetration accordingly.
- **Runway:** cash runway months = cash / quarterly burn; link to financing risk.

---

## 6) Modelling hygiene & **Checks** (must pass)
- Balance sheet balances (Assets = Liabilities + Equity) each period.  
- No #REF! / #DIV/0! / circulars (unless intentional and stable, disclosed).  
- Link integrity (no hardcodes in formulas for key lines).  
- PoS weights sum sanity (0–100% as appropriate to each asset; scenario toggles applied consistently).  
- EV bridge (rNPV/DCF → EV → Equity Value → Target Price) reconciles.  
- Sign conventions consistent (outflows negative).  
- Scenario flags drive the right inputs (Base/Bull/Bear).

**Export for web:** `Checks.csv` with columns **Test, Result, Notes**; “Result” must read **OK** in green in Excel.

---

## 7) Sensitivity policy
Provide at least one **two‑way table** and two **one‑way**:
- Two‑way: **Discount/WACC** vs **Peak sales** (or **PoS**).  
- One‑way: **PoS ±5–10pp**; **Discount/WACC ±1pp**; for M&A: **Synergy** vs **Cost to achieve**; for LBO: **Exit multiple** vs **Net leverage at entry**.
Include spider charts only if they add clarity; otherwise keep tables.

---

## 8) Comps & precedents (sanity checks)
- **Peer set:** 5–8 names, same stage/indication where possible; exclude outliers or disclose.  
- **Metrics:** EV/Revenue, EV/EBITDA (if meaningful), growth‑adjusted where appropriate.  
- **Precedents:** deal value, EV/sales, indication fit, stage, geography.
Keep a short rationale line for each inclusion/exclusion.

---

## 9) Documentation standards
- Every model has a **File Map** tab (what each tab does).  
- **Assumptions tab:** parameter, value, source, date.  
- **Versioning:** `Company_Doc_v1.0_YYYY-MM-DD.ext` with a one‑line “what changed” in the pitch folder README.  
- **CSV exports:** place next to the model inside `models/csv/` so GitHub renders them in‑browser.

---

## 10) Sources & citation convention
- **Regulatory/clinical:** RNS, FDA/EMA releases, ClinicalTrials.gov / EU CTR.  
- **Company:** Annual Report/10‑K, investor presentations (date‑stamped).  
- **Secondary:** reputable industry reports; avoid blogs/rumours.
Footnote tables with short citations (e.g., “AR 2024 p. 115”, “RNS 2025‑03‑14”).

---

## 11) Compliance notes
Public information only; no MNPI; no client/insider data. Forecasts are working estimates and may be wrong. This repository is not investment advice.
