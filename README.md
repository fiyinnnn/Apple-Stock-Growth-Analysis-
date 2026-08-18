# 🍎 Apple Stock Analysis Dashboard

## 📌 Project Overview
<img width="1280" height="683" alt="image" src="https://github.com/user-attachments/assets/a5ce0d24-57d0-4377-b841-c5d889d4c39f" />

This project analyzes **Apple Inc. stock performance** using historical stock-market data in Microsoft Excel.

The dashboard examines:

- 📈 Closing price performance
- 📊 Annual growth rate
- 📉 Moving averages
- 📦 Trading volume
- ⚡ Volatility
- 💰 Overall stock growth
- 📌 Key performance indicators (KPIs)

The analysis transforms raw Apple stock data into calculations, summary tables, charts, and an interactive dashboard.

---

## 📚 Workbook Structure

```text
Apple_Stock_Introduction
Raw_Stock_Data
KPI_Calculations
Closing_Price_Trend
Annual_Growth_Rate
Moving_Average_vs_Closing_Price
Volume_Trend
Volatility
Dashboard
````

---

# 🗂️ 1. Prepare the Raw Stock Data

## Step 1: Create the Raw Data Sheet

Create a worksheet containing:

| Column | Description                             |
| ------ | --------------------------------------- |
| Date   | Trading date                            |
| Open   | Opening stock price                     |
| High   | Highest price during the trading period |
| Low    | Lowest price during the trading period  |
| Close  | Closing stock price                     |
| Volume | Number of shares traded                 |

## Step 2: Convert the Data into an Excel Table

1. Select the complete dataset.
2. Press **Ctrl + T**.
3. Select **My table has headers**.
4. Name the table:

```text
applerevenue
```

## Step 3: Add Supporting Columns

Add:

```text
Profit or Not
Volatility
```

### Volatility Formula

```excel
=High-Low
```

Structured-reference version:

```excel
=applerevenue[[#This Row],[high]]-applerevenue[[#This Row],[low]]
```

<img width="593" height="583" alt="image" src="https://github.com/user-attachments/assets/5a09a820-fa5b-4b60-9904-fe41007c70c8" />


---

# 📊 2. Create the KPI Calculations

Create a worksheet called:

```text
KPI'S
```

The KPI section summarizes the most important results from the analysis.

## KPI 1 — First Closing Price

```excel
=INDEX(applerevenue!E:E,MATCH(MIN(applerevenue!A:A),applerevenue!A:A,0))
```

This retrieves Apple's closing price on the earliest date in the dataset.

---

## KPI 2 — Last Closing Price

```excel
=INDEX(applerevenue!E:E,MATCH(MAX(applerevenue!A:A),applerevenue!A:A,0))
```

This retrieves Apple's latest closing price.

---

## KPI 3 — Overall Growth Rate

```excel
=(Last Close-First Close)/First Close
```

Example:

```excel
=(B4-B3)/B3
```

Format the result as a percentage.

---

## KPI 4 — CAGR

CAGR measures the average annual growth rate over the entire period.

### Calculate Number of Years

```excel
=(MAX(applerevenue!A:A)-MIN(applerevenue!A:A))/365
```

### Calculate CAGR

```excel
=(Last Close/First Close)^(1/Number of Years)-1
```

Example:

```excel
=(B4/B3)^(1/B6)-1
```

---

## KPI 5 — Average Annual Return

```excel
=AVERAGE('Annual Growth Rate'!F4:F45)
```

---

## KPI 6 — Best Growth Year

Highest annual growth:

```excel
=MAX('Annual Growth Rate'!F4:F45)
```

<img width="469" height="167" alt="image" src="https://github.com/user-attachments/assets/02783efb-1ee5-4df5-8f80-ad8978d8affe" />

---

# 📈 3. Closing Price Trend

Create a worksheet called:

```text
Clsing Price Trend
```

## Step 1: Create a PivotTable

Use:

```text
Rows → Date/Year
Values → Close
```

Change the value calculation to:

```text
Average
```

This produces Apple's average closing price for each year.

## Step 2: Create the Chart

1. Select the yearly closing-price summary.
2. Go to **Insert → Line Chart**.
3. Add the title:

```text
Apple Closing Price Trend Over the Years
```

### Recommended Axis

```text
X-axis → Year
Y-axis → Average Closing Price ($)
```

<img width="222" height="413" alt="image" src="https://github.com/user-attachments/assets/93b39dc6-439c-4936-bf4d-59843964b851" />

<img width="510" height="310" alt="image" src="https://github.com/user-attachments/assets/12c08aba-d1ee-4299-ab25-74a42c7db49c" />

---

# 📊 4. Annual Growth Rate

Create a worksheet called:

```text
Annual Growth Rate
```

## Step 1: Calculate Annual Growth

Use:

```excel
=(Current Year Close-Previous Year Close)/Previous Year Close
```

Example:

```excel
=(E5-E4)/E4
```

Format the result as a percentage.

---

## Step 2: Calculate Average Growth by Period

The data can be grouped into five-year periods:

```text
1980-1984
1985-1989
1990-1994
1995-1999
2000-2004
...
```

Calculate the average using:

```excel
=AVERAGE(F4:F8)
```

---

## Step 3: Create the Chart

1. Select the annual growth-rate data.
2. Go to **Insert → Bar Chart**.
3. Title the chart:

```text
Apple Annual Stock Growth Rate (%)
```

4. Format the values as percentages.

<img width="473" height="460" alt="image" src="https://github.com/user-attachments/assets/ef038520-166d-4ab1-a45b-fbed38916035" />

<img width="502" height="301" alt="image" src="https://github.com/user-attachments/assets/82944958-c860-43c6-9cc4-09b7c51dd4bd" />


## 
# 📉 5. Moving Average vs Closing Price

The closing prices were grouped into **5-year periods**, and the average closing price was calculated for each period.

### Step 1: Calculate 5-Year Averages

```excel
=AVERAGE(closing prices for the 5-year period)
````

The resulting 5-year averages were used as the values for the chart.

<img width="473" height="497" alt="image" src="https://github.com/user-attachments/assets/96060940-387e-43f0-97e9-4c7b5b28015e" />

### Step 2: Create the Chart

1. Select the **5-Year Period** and **Average Closing Price** columns.
2. Go to **Insert → Line Chart**.
3. Add the title:
   **Apple 5-Year Average Closing Price Trend**

### Step 3: Add the Trendline

A **3-year trendline** was added to the chart to show the overall direction of the 5-year average closing prices.

1. Click the chart.
2. Select **Chart Elements (+) → Trendline → More Options**.
3. Set the trendline to **3 years**.

<img width="477" height="299" alt="image" src="https://github.com/user-attachments/assets/e3b2e08c-3b26-431b-8d9d-e5b8b7c8fda6" />

```

# 📦 6. Volume Trend

Create a worksheet called:

```text
volume Trend
```

## Step 1: Create a PivotTable

Use:

```text
Rows → Year
Values → Volume
```

Set Volume to:

```text
Sum
```

This calculates Apple's total trading volume for each year.

## Step 2: Create the Chart

1. Select the yearly volume summary.
2. Go to **Insert → Column/Bar Chart**.
3. Title the chart:

```text
Apple Volume Growth Over Time
```

Format large values using:

```text
Millions
Billions
```

where appropriate.

<img width="501" height="534" alt="image" src="https://github.com/user-attachments/assets/4cfef676-4a90-40dd-9ba5-41a0358ec3d5" />

<img width="521" height="359" alt="image" src="https://github.com/user-attachments/assets/d251a337-dbee-4181-8233-41f9f0e792af" />

---

# ⚡ 7. Volatility Analysis

Create a worksheet called:

```text
volatility
```

## Step 1: Calculate Daily Volatility

```excel
=High-Low
```

Structured-reference version:

```excel
=applerevenue[[#This Row],[high]]-applerevenue[[#This Row],[low]]
```

---

## Step 2: Summarize Volatility by Year

Create a PivotTable using:

```text
Rows → Year
Values → Volatility
```

Set the calculation to:

```text
Average
```

---

## Step 3: Create the Chart

1. Select the yearly volatility data.
2. Insert a **Line Chart**.
3. Add the title:

```text
Apple Volatility Rate Over Time
```

<img width="343" height="530" alt="image" src="https://github.com/user-attachments/assets/73f3f289-8f35-406a-9d77-e79fcd17a4c6" />

<img width="237" height="541" alt="image" src="https://github.com/user-attachments/assets/6d59c943-88e4-4959-962a-db949141790b" />

<img width="530" height="325" alt="image" src="https://github.com/user-attachments/assets/92ddbd67-1178-4a42-b7b9-9f9c1ba89fdc" />

---

# 🎛️ 8. Build the Dashboard

Create a worksheet called:

```text
Dashboard
```

The dashboard combines the major findings into one visual report.

## Dashboard Charts

```text
Apple Closing Price Trend Over the Years
Apple Closing Price vs Moving Average
Apple Annual Stock Growth Rate (%)
Apple Volume Growth Over Time
Apple Volatility Rate Over Time
```

---

## Step 1: Create the Dashboard

1. Add a new worksheet.
2. Rename it:

```text
Dashboard
```

3. Hide gridlines.
4. Add the title:

```text
APPLE STOCK ANALYSIS DASHBOARD
```

---

## Step 2: Add KPI Cards

Place the major KPIs at the top:

```text
[ First Close ] [ Last Close ] [ Growth Rate ] [ CAGR ]
[ Avg Return ] [ Best Growth Year ]
```

Use large, readable numbers.

<img width="1182" height="85" alt="image" src="https://github.com/user-attachments/assets/c2fb7521-6265-416d-978c-abd58985aaf1" />

---

## Step 3: Add the Charts

Suggested layout:

```text
┌───────────────────────────────┬───────────────────────────────┐
│ Closing Price Trend            │ Moving Average               │
│                                │                              │
├───────────────────────────────┼───────────────────────────────┤
│ Annual Growth Rate             │ Volume Trend                 │
│                                │                              │
├───────────────────────────────┴───────────────────────────────┤
│                    Volatility Trend                            │
└────────────────────────────────────────────────────────────────┘
```


<i<img width="1280" height="683" alt="image" src="https://github.com/user-attachments/assets/c22e3fe0-8990-435a-b635-3012137bb4e9" />


---

# 🎨 9. Dashboard Formatting

## Number Formatting

### Stock Prices

```text
$0.00
```

### Growth Rates

```text
0.00%
```

### Large Volume Values

```text
0.00M
```

or:

```text
0.00B
```

---

## Layout Guidelines

* Keep equal spacing between charts.
* Align charts consistently.
* Remove unnecessary borders.
* Use clear chart titles.
* Keep KPI values larger than supporting text.
* Avoid overcrowding the dashboard.
* Maintain a consistent visual theme.

---

# 🔍 10. Questions the Dashboard Answers

## Price Performance

* How has Apple's closing price changed over time?
* Which periods experienced the strongest price growth?

## Growth

* What was Apple's annual growth rate?
* Which year recorded the highest growth?
* What was Apple's overall growth?
* What is Apple's CAGR?

## Trend

* How does the moving average compare with the closing price?
* When did the stock price move significantly from its moving average?

## Trading Activity

* How has Apple's trading volume changed?
* Which periods recorded the highest trading volume?

## Risk

* How has Apple's volatility changed?
* Which periods experienced greater price fluctuations?

---

# 🧹 11. Final Quality Checklist

* [ ] Raw data contains the correct headers.
* [ ] Dates are formatted correctly.
* [ ] Closing prices are formatted as currency.
* [ ] Growth rates are formatted as percentages.
* [ ] Volume uses readable units.
* [ ] KPI formulas return the correct values.
* [ ] Charts have clear titles.
* [ ] Axis labels are understandable.
* [ ] Moving-average calculations are correct.
* [ ] Volatility calculations are correct.
* [ ] Dashboard charts are aligned.
* [ ] Gridlines are removed.
* [ ] Dashboard is easy to read.

---

# 📁 Final Workbook Structure

```text
Apple Stock Dashboard.xlsx
│
├── KPI'S
├── Dashboard
├── Clsing Price Trend
├── Annual Growth Rate
├── MA VS closing proce trend
├── volume Trend
├── volatility
├── applerevenue
├── Sheet6
└── ma vs cpt
```

---

# 🏁 Conclusion

The Apple Stock Analysis Dashboard transforms historical Apple stock data into a visual analysis of:

* **Price performance**
* **Growth**
* **Trends**
* **Trading activity**
* **Volatility**

This project demonstrates practical Excel skills including:

* Data cleaning and organization
* Excel Tables
* PivotTables
* Excel formulas
* KPI development
* Moving averages
* Growth-rate calculations
* CAGR
* Data visualization
* Dashboard design
* Financial data analysis

The final dashboard provides a concise view of Apple's historical stock performance while allowing the calculations to be traced back to the original dataset.

