# 🏠 Home Deposit Investment Strategy Analysis

> A quantitative Excel/VBA model for personal financial planning — blending actuarial science, fixed-income modelling, equity analysis, and stochastic simulation to build a 10-year investment roadmap.

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Background & Motivation](#2-background--motivation)
3. [Key Features](#3-key-features)
4. [Technical Methodology](#4-technical-methodology)
   - 4.1 [Income & Salary Modelling](#41-income--salary-modelling)
   - 4.2 [Savings Framework](#42-savings-framework)
   - 4.3 [Bond Pricing — Nelson-Siegel-Svensson Model](#43-bond-pricing--nelson-siegel-svensson-model)
   - 4.4 [Equity Returns — VAS Historical Analysis](#44-equity-returns--vas-historical-analysis)
   - 4.5 [Risky Asset X — Market-Condition Dependent Simulation](#45-risky-asset-x--market-condition-dependent-simulation)
   - 4.6 [Portfolio Strategies](#46-portfolio-strategies)
   - 4.7 [Stochastic Projection & Scenario Analysis](#47-stochastic-projection--scenario-analysis)
5. [Excel Workbook Structure](#5-excel-workbook-structure)
6. [VBA Automation](#6-vba-automation)
7. [Results & Insights](#7-results--insights)
8. [Skills Demonstrated](#8-skills-demonstrated)
9. [Getting Started](#9-getting-started)
10. [Prerequisites](#10-prerequisites)
11. [Future Extensions](#11-future-extensions)
12. [About](#12-about)

---

## 1. Project Overview

This project develops a comprehensive, data-driven financial planning model for **Jackson** — a recent actuarial graduate beginning his career — who wants to save for a home deposit over a **10-year horizon (2023–2033)**.

The model is built entirely in **Microsoft Excel with VBA macros**, combining:

- **Actuarial income modelling** (salary growth, bonus probabilities)
- **Fixed-income analytics** (yield-curve fitting via the Nelson-Siegel-Svensson model)
- **Equity return simulation** (Australian shares index VAS using bootstrap resampling of historical data)
- **Alternative/risky asset simulation** (market-regime-dependent returns for Asset X)
- **Monte Carlo-style scenario analysis** (projecting future portfolio values under uncertainty)

The output is a rich set of projections and recommendations tailored to Jackson's risk tolerance and financial goals.

---

## 2. Background & Motivation

Saving for a home deposit in Australia's real-estate market is a significant financial challenge, especially for young professionals. The decision of *how* to invest one's savings — across bonds, equities, and higher-risk alternatives — can have a substantial impact on the final accumulated sum.

This project was motivated by the desire to apply rigorous quantitative methods (typically used in institutional finance and actuarial science) to a personal financial planning problem. It demonstrates that the same models used in professional investment management are accessible at the individual level, given the right analytical toolkit.

The project also serves as an exploration of:
- How different **asset allocation strategies** perform over a medium-term horizon under varying market conditions
- The impact of **income uncertainty** (bonus variability) on long-term savings
- How **downside risk** in a risky asset affects the overall portfolio

---

## 3. Key Features

| Feature | Description |
|---|---|
| **Dynamic Income Projections** | Month-by-month salary modelling with contractual increments, sign-on bonus, and probabilistic performance bonuses |
| **Flexible Savings Rate** | Savings calculated as a configurable percentage of total monthly income |
| **Nelson-Siegel-Svensson Bond Returns** | Yield-curve fitting for Australian government bonds to project bond investment returns |
| **VAS Equity Simulation** | Bootstrap resampling of historical VAS (Vanguard Australian Shares ETF) return data to simulate future equity performance |
| **Regime-Dependent Asset X** | Market-condition-based simulation for a risky alternative asset with distinct bull, neutral, and bear market return distributions |
| **Three Distinct Portfolio Strategies** | Comparison of conservative, balanced, and aggressive asset allocation approaches |
| **10-Year Savings Projection** | Future value calculations under multiple scenarios, with compounding reinvestment |
| **Risk & Downside Analysis** | Quantification of the impact of adverse market conditions on Portfolio outcomes |
| **Summary Dashboard** | Consolidated sheet with key metrics, scenario comparisons, and tailored recommendations |
| **VBA Macros** | Automated recalculation, scenario switching, and data refreshing via VBA |

---

## 4. Technical Methodology

### 4.1 Income & Salary Modelling

Jackson's monthly income is modelled with the following components:

**Base Salary:**
- Starts at a graduate level and grows at contractual annual increment rates
- Monthly salary = Annual salary ÷ 12

**Sign-On Bonus:**
- A one-time lump sum received at the start of employment
- Directly added to the initial savings pool

**Performance Bonus:**
- Modelled as a probabilistic event: each year, Jackson receives a bonus with a specified probability
- Bonus magnitude is expressed as a percentage of annual salary, itself drawn from a defined distribution
- Expected bonus = P(bonus) × E[bonus % | bonus occurs] × Annual Salary

This probabilistic treatment of bonuses introduces stochasticity into the savings trajectory, more accurately reflecting real-world income uncertainty.

---

### 4.2 Savings Framework

Each month, Jackson saves a fixed percentage **s** of his total income (base + bonuses). The savings rate **s** is a model parameter that can be adjusted to test sensitivity.

```
Monthly Savings(t) = s × [Base Salary(t) + Bonus(t)]
```

Accumulated savings are then invested across the three strategies, with returns compounding monthly.

---

### 4.3 Bond Pricing — Nelson-Siegel-Svensson Model

Bond returns are derived by fitting the **Nelson-Siegel-Svensson (NSS)** yield curve to observed Australian government bond yield data. The NSS model parameterises the yield curve as:

```
y(τ) = β₀
     + β₁ · [(1 - exp(-τ/λ₁)) / (τ/λ₁)]
     + β₂ · [(1 - exp(-τ/λ₁)) / (τ/λ₁) - exp(-τ/λ₁)]
     + β₃ · [(1 - exp(-τ/λ₂)) / (τ/λ₂) - exp(-τ/λ₂)]
```

Where:
- `τ` is the time to maturity
- `β₀` is the long-term yield level
- `β₁` controls the short-term component
- `β₂` controls the medium-term hump/trough
- `β₃` is an additional curvature term
- `λ₁`, `λ₂` are decay rate parameters

Parameters are calibrated to observed yield data using a least-squares optimisation. The fitted curve is then used to:
1. Price zero-coupon bonds at any maturity
2. Calculate the annual return on a bond portfolio over the investment horizon
3. Project bond portfolio value year-by-year

This approach captures the full shape of the yield curve — including inversions — rather than relying on a single yield figure.

---

### 4.4 Equity Returns — VAS Historical Analysis

For the equities component, historical return data from the **Vanguard Australian Shares ETF (VAS)** is used to model future equity performance.

The methodology involves:

1. **Historical return series**: Monthly total returns (price appreciation + dividends) from VAS over a multi-year lookback window are collected
2. **Descriptive statistics**: Mean, standard deviation, skewness, and kurtosis of monthly returns are computed to characterise the distribution
3. **Bootstrap resampling**: Rather than assuming a parametric distribution (e.g. normal), future return paths are simulated by sampling with replacement from the historical monthly return series — preserving the empirical distribution, including fat tails and non-normality
4. **Annual return calculation**: Compounded from 12 sampled monthly returns

This non-parametric approach avoids the assumption of normally distributed returns and better reflects the actual statistical properties of equity returns (negative skewness, excess kurtosis).

---

### 4.5 Risky Asset X — Market-Condition Dependent Simulation

Asset X represents a higher-risk, higher-potential-return alternative investment (e.g. a start-up, speculative asset, or high-yield instrument).

Its return model is **regime-dependent**, distinguishing three market states:

| Market Condition | Probability | Expected Return | Characteristics |
|---|---|---|---|
| **Bull Market** | p₁ | μ_bull (high) | Strong positive returns, low volatility |
| **Neutral Market** | p₂ | μ_neutral (moderate) | Returns near long-run average |
| **Bear Market** | p₃ | μ_bear (negative) | Losses possible; captures downside risk |

Where p₁ + p₂ + p₃ = 1.

The annual return for Asset X is sampled as:
1. Draw the market state from the categorical distribution {Bull, Neutral, Bear}
2. Sample the return from the state-specific distribution (modelled as normal with state-specific mean and variance)

**Downside Risk Analysis:**
A specific scenario analysis quantifies the impact if Asset X enters an extended bear phase — simulating multi-year negative return regimes and their effect on the overall portfolio. This stress-testing is critical for understanding the portfolio's vulnerability to tail risks.

---

### 4.6 Portfolio Strategies

Three distinct strategies are evaluated, each representing a different risk-return profile:

| Strategy | Bonds | VAS Equities | Risky Asset X | Target Investor |
|---|---|---|---|---|
| **Strategy 1 (Conservative)** | High allocation | Moderate allocation | None or minimal | Risk-averse; capital preservation priority |
| **Strategy 2 (Balanced)** | Moderate allocation | High allocation | Small allocation | Moderate risk tolerance; growth-oriented |
| **Strategy 3 (Aggressive)** | Low allocation | Moderate allocation | High allocation | High risk tolerance; maximising expected return |

Each strategy's projected value at the 10-year mark is computed under multiple market scenarios, enabling a direct comparison of risk-adjusted outcomes.

---

### 4.7 Stochastic Projection & Scenario Analysis

The model projects the accumulated portfolio value over 120 months (10 years) under the following scenarios:

- **Base Case**: Expected returns for all assets; expected bonus receipts
- **Optimistic Case**: Bull market conditions; maximum bonuses
- **Pessimistic Case**: Bear market for Asset X; minimal bonuses; lower equity returns
- **Downside Stress Test**: Extended negative returns for Asset X; no bonuses

For each scenario and strategy, the model computes:
- **Monthly portfolio value** (contributions + reinvested returns)
- **Final accumulated savings** at month 120
- **Probability of reaching the home deposit target** (based on scenario outcomes)
- **Shortfall/surplus** relative to the deposit goal

The comparison across strategies and scenarios enables a **risk-return tradeoff analysis**, directly informing the recommendation made to Jackson.

---

## 5. Excel Workbook Structure

The file `Excel-VBA-Investment-Analysis.xlsm` contains the following sheets:

| Sheet | Contents |
|---|---|
| **Income & Investment** | Month-by-month breakdown of Jackson's salary, bonuses, savings contributions, and total funds available for investment |
| **Investment Options** | Detailed calculations for each of the three investment strategies, including asset allocations and expected returns per period |
| **Bond Analysis** | NSS model parameter fitting, yield curve outputs, and bond portfolio return projections |
| **VAS Equity Data** | Historical VAS return data, descriptive statistics, and bootstrap simulation results |
| **Risky Asset X** | Market regime probability inputs, state-specific return distributions, and simulated annual returns with scenario toggles |
| **Summary & Recommendations** | Dashboard consolidating key results: projected final balances, scenario comparisons, risk metrics, and strategy recommendations tailored to Jackson's profile |

---

## 6. VBA Automation

The workbook includes VBA macros to automate key tasks:

- **Scenario Switcher**: A macro-driven control that updates all linked formulas and projection tables when switching between Base, Optimistic, Pessimistic, and Stress scenarios — avoiding manual cell-by-cell updates
- **Simulation Re-run**: Triggers a fresh bootstrap resample of VAS returns and a new draw from the Asset X regime model, allowing the user to explore the distribution of outcomes across multiple simulation runs
- **Report Generator**: Populates the Summary sheet with the latest computed values from across all analysis sheets, ensuring consistency
- **Input Validation**: Checks that user-entered parameters (savings rate, allocation percentages, bonus probabilities) are within valid ranges before running projections

These macros reduce the risk of manual errors and make the model more accessible to non-technical users.

---

## 7. Results & Insights

Key findings from the analysis (based on the model's base-case parameterisation):

- **Strategy 1 (Conservative)** provided the most predictable savings trajectory, with the lowest variance in final accumulated balance. However, its expected terminal value was the lowest of the three strategies, potentially falling short of the home deposit target in low-return environments.

- **Strategy 2 (Balanced)** offered a compelling risk-return tradeoff — achieving a meaningfully higher expected terminal value than the conservative strategy, while keeping downside risk (bear market scenario outcomes) manageable.

- **Strategy 3 (Aggressive)** showed the highest expected terminal value but also the widest spread of outcomes. In the bear market stress-test scenario, significant drawdowns in Asset X substantially eroded the portfolio, highlighting the concentration risk of high alternative-asset exposure.

- **Income uncertainty** (bonus variability) had a notable compounding effect over 10 years — the difference between "no bonuses received" and "maximum bonuses received" scenarios represented a substantial portion of the final deposit target.

- The **Nelson-Siegel-Svensson bond model** captured the current low-yield environment accurately and revealed that bond-heavy portfolios would struggle to keep pace with the target deposit growth rate in the medium term.

- **Recommendation for Jackson**: A Balanced portfolio (Strategy 2) aligned best with his profile as a young professional with a stable primary income, moderate risk tolerance, and a defined 10-year investment horizon.

---

## 8. Skills Demonstrated

This project showcases a range of quantitative and technical skills directly relevant to data science, ML engineering, and quantitative finance roles:

| Skill Area | Specific Application |
|---|---|
| **Mathematical Modelling** | Nelson-Siegel-Svensson yield curve fitting; probabilistic income modelling |
| **Statistical Analysis** | Descriptive statistics of financial return distributions; non-parametric bootstrap resampling |
| **Stochastic Simulation** | Regime-switching models for risky asset returns; scenario generation |
| **Optimisation** | Least-squares parameter calibration for the NSS model |
| **Data Analysis** | Processing and interpreting historical time-series return data (VAS ETF) |
| **Risk Quantification** | Downside risk analysis, stress testing, scenario-based sensitivity analysis |
| **Financial Engineering** | Bond pricing, yield curve construction, portfolio compounding, future value projection |
| **Spreadsheet Engineering** | Complex, structured Excel model with linked sheets, dynamic formulas, and data validation |
| **VBA / Automation** | Macro-driven scenario switching, simulation re-runs, and report generation |
| **Communication** | Translating complex model outputs into clear, actionable recommendations for a non-expert audience |

---

## 9. Getting Started

1. **Clone or download** this repository
2. Open `Excel-VBA-Investment-Analysis.xlsm` in **Microsoft Excel** (2016 or later recommended)
3. If prompted, **enable macros** (required for VBA automation features)
4. Navigate to the **Summary & Recommendations** sheet for an overview of results
5. To explore the model inputs, open the **Income & Investment** sheet and adjust parameters (highlighted in yellow) such as:
   - Savings rate (%)
   - Bonus probability and magnitude assumptions
   - Asset allocation percentages per strategy
6. Use the **Scenario Switcher** macro button to toggle between scenarios and observe the updated projections

> **Note:** Some sheets contain protected formulas to prevent accidental edits. The model's input cells (yellow-highlighted) are intentionally left unprotected for user customisation.

---

## 10. Prerequisites

- **Microsoft Excel 2016+** (Windows or Mac)
  - Macro execution must be enabled (`.xlsm` format)
  - The **Analysis ToolPak** add-in is recommended for extended statistical functions
- **No external data connections required** — all data is embedded within the workbook
- **Basic familiarity with Excel** is sufficient to navigate and use the model; understanding of fixed-income or equity concepts is helpful for interpreting results

---

## 11. Future Extensions

Several enhancements could make this model more powerful and bring it closer to a production-grade financial planning tool:

- **Monte Carlo Simulation at Scale**: Replace the scenario-based approach with a full Monte Carlo engine (thousands of simulated paths) to produce proper probability distributions over terminal portfolio values — this is a natural next step towards implementing this in Python with `NumPy`/`pandas`
- **Python/Jupyter Reimplementation**: Port the model to Python for reproducibility, version control compatibility, and integration with real-time data APIs (e.g. ASX data via `yfinance`)
- **Machine Learning for Return Prediction**: Explore whether ML models (e.g. gradient boosting, LSTM networks) trained on macroeconomic features can improve upon the regime-switching model for Asset X returns
- **Dynamic Asset Allocation (Rebalancing)**: Incorporate periodic portfolio rebalancing, where allocations shift as market conditions change or as Jackson approaches his 10-year target
- **Tax Modelling**: Add Australian income tax, capital gains tax (CGT), and franking credit calculations to produce after-tax return estimates
- **Inflation Adjustment**: Adjust the home deposit target and all cash flows for CPI inflation to produce real (inflation-adjusted) projections
- **Goal-Based Optimisation**: Formulate the asset allocation as an optimisation problem — maximise the probability of reaching the deposit target subject to a maximum acceptable drawdown constraint
- **Interactive Dashboard**: Build a Streamlit or Dash web app front-end to make the model accessible via a browser without requiring Excel

---

## 12. About

This project was developed as a personal passion project exploring the intersection of **actuarial science**, **quantitative finance**, and **data-driven decision making**. It represents the kind of rigorous, model-first thinking that underpins work in ML engineering, AI systems design, and quantitative research.

It is featured in the *Passion Projects & Interests* section of my portfolio to demonstrate:
- Comfort with mathematical and statistical modelling from first principles
- Ability to build end-to-end analytical systems (data → model → insight → recommendation)
- Interest in applying computational methods to real-world problems outside of a formal work setting

---

*Built with Microsoft Excel & VBA · Quantitative Methods: Nelson-Siegel-Svensson, Bootstrap Resampling, Regime-Switching Models*

