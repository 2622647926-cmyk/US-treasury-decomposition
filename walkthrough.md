# US Treasury Decomposition Dashboard - Walkthrough

This document summarizes the changes, visual layouts, and verification results of the **US Treasury Decomposition Dashboard** project.

## 🚀 Key Improvements & Accomplishments

### 1. Macro Analytical Attribution Panel (Decision Tree Model)
- **Visualized Thought Flow**: Created a dedicated, interactive HTML/CSS sidebar visualizing the macro decision tree: 名义利率 -> TIPS 实际 vs BEI 通胀 -> 短端 vs 长端 -> ACM 期限溢价 TP 检验。
- **Dynamic Diagnostics Calculator**: The user can choose any start/end date or select 1-month, 3-month, or year-to-date presets. The calculator computes the precise basis point (bp) changes of all variables and automatically generates an institutional-grade narrative report analyzing the dominant fundamental and risk-premium drivers step-by-step.

### 2. Chart 10 Multi-Axis Inverted Yield Alignment
- **3-Y-Axes Plotting**: Updated Chart 10 to support three independent Y-axes.
- **Inverted 2Y Series**: Plotted the 2Y Nominal Yield on a dedicated inverted Y-axis (`inverse: true` with a 55px right-side offset). This visually aligns the negative correlation of interest rates and Gold, showing them moving in tandem.

### 3. Chart 5 Dual-Perspective Z-Scores (Rolling vs Full-Sample)
- **Rolling Z-score (Pink)**: Standard 252-day window Z-score, excellent for detecting short-term regimes and localized anomalies.
- **Full-Sample Z-score (Amber)**: Based on the entire sample since 2000. It doesn't adapt to rolling windows, solving the adaptation issue where Z-scores converged to 0 despite yields setting new historic highs.

### 4. Direct New York Fed ACM Daily Excel Extraction
- Avoided the copyrighted/404 FRED ID constraints by writing a custom scraper that downloads the official 10MB `ACMTermPremium.xls` directly from the New York Fed website.
- Parsed the daily zero-coupon yield, expected short rate, and term premium directly from the `ACM Daily` sheet.

---

## 🛠️ Verification & Test Runs
- The Python script `fetch_data.py` exited with Code 0.
- All 16,261 daily ACM entries have been downloaded and successfully merged with FRED and Yahoo Finance.
- `data.js` has been successfully updated with the new `fwd_deviation_zscore_full` metric.
