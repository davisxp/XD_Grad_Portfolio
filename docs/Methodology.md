# Methodology — Life‑Sciences Finance Portfolio
**Maintainer:** Xavier Davis • **Updated:** 27 Oct 2025  
Purpose: establish consistent, auditable methods for valuation, modelling, and documentation across all six pitches.

---

## Valuation frameworks
- **DCF:** 5–10 year explicit with terminal value (g or exit multiple), WACC baseline 8–10% for diversified/medtech, case‑by‑case for pure‑play biotech with marketed assets.
- **Comps:** 5–8 peers; normalise for growth/margins; triangulate to sanity‑check the DCF/rNPV output.
- **Merger:** sources & uses; purchase price allocation (PPA); synergy timing/cost to achieve; debt schedule; accretion/dilution.
- **LBO:** tranches, fees, cash sweep, covenant headroom, MOIC/IRR with entry/exit bridges.
- **LevFin/DCM:** rating‑case vs base‑case metrics (gross/net leverage, ICR, FCF conversion), maturity wall, covenant capacity.

---

## Probability of Success (PoS) — baseline table
Can adjust by indication, MoA novelty, trial design, prior data quality. These are approximates.

| Stage | General | Oncology |
|---|---:|---:|
| Preclinical | 5–10% | 3–8% |
| Phase I | 10–20% | 8–15% |
| Phase II | 20–40% | 15–30% |
| Phase III | 50–70% | 45–60% |
| Filed/Regulatory | 80–90% | 75–85% |

**Note:** If robust  endpoints or breakthrough pathways apply, tilt upward; if single‑site/small‑N or safety flags, tilt downward.

---

## Revenue drivers & uptake
- **Uptake curve:** logistic/S‑curve; specify time to peak (e.g., 4–6 years).  
- **Pricing:** list price, gross‑to‑net (rebates, access), geography mix if used.  
- **Competition:** show a 2–3 line competitive map and adjust penetration accordingly.
- **Runway:** cash runway months = cash / quarterly burn; link to financing risk.

---

## 6) Modelling hygiene & **Checks** 
- Balance sheet balances (Assets = Liabilities + Equity) each period.  
- No #REF! / #DIV/0! / circulars (unless intentional and stable, disclosed).  
- Link integrity (no hardcodes in formulas for key lines).  
- PoS weights sum sanity (0–100% as appropriate to each asset; scenario toggles applied consistently).  
- EV bridge (rNPV/DCF → EV → Equity Value → Target Price) reconciles.  
- Sign conventions consistent (outflows negative).  
- Scenario flags drive the right inputs (Base/Bull/Bear).

**Export for web:** `Checks.csv` with columns **Test, Result, Notes**; “Result” must read **OK** in green in Excel.

---

## Sensitivity
Provide at least one **two‑way table** and two **one‑way**:
- Two‑way: **Discount/WACC** vs **Peak sales** (or **PoS**).  
- One‑way: **PoS ±5–10pp**; **Discount/WACC ±1pp**; for M&A: **Synergy** vs **Cost to achieve**; for LBO: **Exit multiple** vs **Net leverage at entry**.

---

## Comps & precedents
- **Peer set:** 5–8 names, same stage/indication where possible; exclude outliers or disclose.  
- **Metrics:** EV/Revenue, EV/EBITDA (if meaningful), growth‑adjusted where appropriate.  
- **Precedents:** deal value, EV/sales, indication fit, stage, geography.
Keep a short rationale line for each inclusion/exclusion.

---

## Documentation standards
- Every model has a **File Map** tab (what each tab does).  
- **Assumptions tab:** parameter, value, source, date.  
- **CSV exports:** place next to the model inside `models/csv/` so GitHub renders in‑browser.

---

## Sources & citation convention
- **Regulatory/clinical:** RNS, FDA/EMA releases, ClinicalTrials.gov / EU CTR.  
- **Company:** Annual Report/10‑K, investor presentations (date‑stamped).  
- **Secondary:** reputable industry reports; avoid blogs/rumours.
Footnote tables with short citations (e.g., “AR 2024 p. 115”, “RNS 2025‑03‑14”).

---
