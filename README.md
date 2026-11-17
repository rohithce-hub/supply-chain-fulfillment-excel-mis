# Supply Chain & E-Commerce Order Fulfillment MIS Dashboard (Excel)

[![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://products.office.com/excel)
[![Kaggle Dataset](https://img.shields.io/badge/Kaggle-DataCo_Supply_Chain-20BEFF?style=for-the-badge&logo=kaggle&logoColor=white)](https://www.kaggle.com/datasets/shashwatwork/dataco-smart-supply-chain-for-big-data-analysis)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)

An end-to-end executive **Management Information System (MIS) Dashboard** built in Microsoft Excel analyzing **1,500 e-commerce order shipments** (based on Kaggle's DataCo Global Supply Chain dataset). The dashboard monitors order fulfillment velocity, delivery SLA compliance, freight expenditure, and regional logistics bottlenecks across 10 commercial hubs.

---

## 📌 Executive Dashboard Preview

![Supply Chain Fulfillment MIS Dashboard](screenshots/Supply_Chain_Fulfillment_MIS_Dashboard.png)

> **Live File**: Open [`Supply_Chain_Fulfillment_MIS_Dashboard.xlsx`](Supply_Chain_Fulfillment_MIS_Dashboard.xlsx) in Microsoft Excel to interact with the clickable native Slicers and dynamic Pivot calculations.

---

## 📊 Executive Summary & Key Performance Indicators (Jan – Oct 2026)

```text
======================================================================================================
LOGISTICS HEALTH & SLA PERFORMANCE AT A GLANCE (1,500 ORDERS)
======================================================================================================
Total Orders Processed : 1,500 Orders
On-Time Deliveries     : 1,175 Orders (78.33% Fulfillment Rate)
Delayed Shipments      : 236 Orders (15.73% Delay Rate)
Cancelled / Attrited   : 54 Orders (3.60%)
Pending / In-Transit   : 35 Orders (2.33%)
Total Order Value      : ₹ 88,36,686 (~$106K USD)
======================================================================================================
```

### 🔍 Key Business Findings:
1. **The Fulfillment SLA Gap**: 78.33% of shipments arrived strictly on schedule, leaving **15.73% (236 shipments)** delayed. Root-cause analysis indicates standard road freight experiences peak delays during multi-state hub transits.
2. **Category Demand Volume**: **Apparel (484 orders)** and **Footwear (359 orders)** represent **56.2% of total order volume**, requiring dedicated buffer stock and prioritized carrier allocation.
3. **Logistics Lead-Time Variance**:
   - **Same Day**: Scheduled 0.0 days vs. Actual 0.1 days (Near-perfect 98% SLA compliance).
   - **First Class**: Scheduled 1.5 days vs. Actual 1.6 days (Minimal variance).
   - **Second Class**: Scheduled 2.5 days vs. Actual 2.6 days.
   - **Standard Class**: Scheduled 4.5 days vs. Actual 4.4 days (Largest absolute volume of delayed shipments).
4. **Regional Hub Concentration**: **Bengaluru (300 orders)**, **Kochi (260 orders)**, and **Mumbai (214 orders)** account for over **51.6% of total dispatched freight**, identifying South and West distribution corridors as critical operational arteries.

---

## 🛠️ Dashboard Architecture & Visual Design

The dashboard follows modern enterprise MIS design principles with zero gridline clutter, high-contrast visual hierarchy, and instant drill-down capabilities:

### 1. Interactive Slicer Panel
* **Fulfillment Status Slicer**: Filter by *Delivered On-Time*, *Delayed*, *Cancelled*, and *Pending*.
* **Shipping Mode Slicer**: Multi-select across *Standard Class*, *Second Class*, *First Class*, and *Same Day*.
* **Warehouse Hub Slicer**: Real-time drill down into *North Hub (Delhi)*, *South Hub (Bengaluru)*, *West Hub (Mumbai)*, and *East Hub (Kolkata)*.

### 2. Dual-Series Category Clustered Column Chart
* Compares **Total Order Volume (Navy Blue)** against **On-Time Deliveries (Emerald Green)** across *Apparel*, *Footwear*, *Outdoors*, *Sporting Goods*, and *Technology*, instantly highlighting fulfillment efficacy per merchandise line.

### 3. Fulfillment Status Donut Distribution
* Clear proportional breakdown with leader lines:
  * 🟢 **78% Delivered On-Time**
  * 🔴 **16% Delayed Shipments**
  * 🟡 **4% Cancelled Orders**
  * 🟣 **2% Pending / In-Transit**

### 4. Dual-Metric Monthly Trend Line (Jan – Oct 2026)
* Visualizes dispatch velocity from **139 orders in January** to a peak of **173 orders in October** with distinct marker geometries and independent data label positioning (above for total volume, below for on-time deliveries).

### 5. Delivery SLA Lead Time vs. Scheduled Commitment
* Clustered column visual tracking scheduled delivery commitments against real-world transit days across all shipping tiers.

### 6. Top Destination Delivery Hubs
* Horizontal bar chart ranking the top 7 metro destination hubs to guide carrier contract renegotiations and hub expansion.

---

## 🧩 Technical Implementation & Excel Engineering

The project implements a modern **3-Sheet Model Architecture** leveraging native Excel **PivotTables**, **PivotCharts**, and synchronized **Slicer Report Connections** with dynamic `=GETPIVOTDATA(...)` formulas:

```mermaid
graph TD
    A["Raw_Data Sheet<br/>(1,500-Row Official OrdersTable)"] -->|Master PivotCache Connection| B["Pivot_Calcs Sheet<br/>(6 Dedicated Analytical PivotTables)"]
    B -->|Dynamic GETPIVOTDATA Formulas| C["KPI Cards Engine<br/>(Real-Time Recalculation)"]
    B -->|Native PivotChart Source Data| D["Dashboard Sheet<br/>(5 Synchronized PivotCharts)"]
    E["Interactive Slicers Panel<br/>(Status, Mode, Warehouse Hub)"] -->|SlicerCache Report Connections| B
    E -.->|Instant Multi-Chart Dynamic Filtering| D
    E -.->|Instant Metric Recalculation| C
```

### Advanced Excel Engineering Features:
* **Native Multi-Pivot Slicer Synchronization (Report Connections)**:
  Three interactive Slicers (`Fulfillment_Status`, `Shipping_Mode`, `Warehouse_Hub`) are bound via programmatic **Report Connections** across all 6 underlying PivotTables. Selecting or deselecting any slicer item dynamically and instantaneously re-filters all 5 charts and KPI cards simultaneously.
* **Dynamic KPI Metric Engine (`GETPIVOTDATA`)**:
  Rather than static range formulas, executive KPI summary cards evaluate the visible PivotTable cache via:
  ```excel
  =GETPIVOTDATA("Orders_Count", Pivot_Calcs!$B$2)
  =GETPIVOTDATA("OnTime_Sum", Pivot_Calcs!$B$2)
  =GETPIVOTDATA("Delayed_Sum", Pivot_Calcs!$B$2)
  =IFERROR(C11/C10, 0)
  =GETPIVOTDATA("Total_Revenue", Pivot_Calcs!$B$2)
  ```
  This guarantees 100% mathematical accuracy and zero `#REF!` errors under any combination of slicer drill-downs.
* **Automated Data Normalization & Flagging**:
  Helper boolean flags (`Is_On_Time`, `Is_Delayed`) embedded in the transactional data table enable dual-metric summations within a single unified PivotTable.
* **Field Button Suppression & Publication Cleanliness**:
  All native PivotChart field buttons and filter dropdown pills are suppressed (`ShowAllFieldButtons = False`) to maintain a clean, publication-ready executive aesthetic.

---

## 🚀 Business Recommendations & Actionable Strategy

Based on the 10-month operational data:
1. **Standard Class Carrier Re-negotiation**: Standard class accounts for over 55% of volume and 88% of delayed orders. Introduce penalty-backed SLA guarantees with 3PL courier partners for delays exceeding 5 transit days.
2. **South Corridor Forward-Deployment**: With Bengaluru and Kochi driving over 560 orders, establish micro-fulfillment hubs in South India to cut transit lead time from 2.6 days to sub-24 hours.
3. **Automated Delay Exception Alerts**: Trigger real-time customer notification for orders flagged with `Actual_Days > Scheduled_Days` by Day 3 to proactively curb customer cancellations.

---

## 📂 Repository Contents

```text
├── Supply_Chain_Fulfillment_MIS_Dashboard.xlsx   # Main Excel file with Data, Calcs & Dashboard
├── Supply_Chain_Fulfillment_MIS_Dashboard.pdf    # 1-Page Vector PDF export of the Dashboard
├── Supply_Chain_Fulfillment_MIS_Dashboard.png    # High-resolution dashboard image for presentations
├── DataCoSupplyChainDataset.csv                  # Original Kaggle DataCo source dataset
├── DescriptionDataCoSupplyChain.csv              # Metadata dictionary & feature descriptions
└── README.md                                     # Project documentation & business insights
```

---

## 👤 Author & Contact

**Rohith Chandran E**  
*Aspiring Data Analyst & MIS Professional*  
* **LinkedIn**: [linkedin.com/in/rohithchandrane](https://www.linkedin.com/in/rohithchandrane)  
* **GitHub**: [github.com/rohithce-hub](https://github.com/rohithce-hub)  
* **Email**: rohith.e.chandran@gmail.com
