# 🍪 B&J Biscuit Business Analysis Dashboard

## 📘 Project Overview
This project presents an **interactive Excel dashboard** designed to analyze key business performance indicators for **B&J Biscuit**.  
The dashboard provides insights into **revenue distribution**, **profitability**, **customer behavior**, and **geographic performance** — enabling quick, data-driven decision-making.

---

## 🎯 Objectives
- Provide a **comprehensive view** of sales and revenue performance  
- Identify **profitable brands, locations, and customer segments**  
- Analyze **customer demographics** and **payment behaviors**  
- Track **Month-over-Month (MoM)** and **Week-over-Week (WoW)** trends  
- Enable **interactive filtering** for flexible analysis  

---

## 📊 Dashboard 1 – Business Performance Overview

### 🔹 Key Metrics
- **Revenue Distribution**
  - By product price category (**Low ≥ 10**, **High < 10**)  
  - By **age group** and **gender**  
  - By **payment method**  

- **Profitability Analysis**
  - Most profitable **brand**, **location**, **customer**, and **salesperson**  
  - Overall **profit margin**

- **Customer Insights**
  - Top 5 customers by **revenue contribution**  
  - Total **number of customers acquired**

- **Geographic Revenue Distribution**
  - Revenue share by **key locations**

- **Sales Performance**
  - Quantity sold, total COGS, total revenue, and total profit  

### 🧩 Interactive Features
- Filters for **Location**, **Payment Method**, and **Age Group**  
- “**Clear Slicers**” button powered by **VBA macros**  
- Clean, user-friendly layout with clear visual storytelling  

---

## 🧮 VBA Automation – Clear Slicers

To improve user experience, I added a **VBA macro** that resets all slicers to their default state with a single click.  
This saves time when exploring different filters or preparing the dashboard for presentations.

### 🔧 Example Code Snippet:
```vba
Sub ClearSlicers()
    With ActiveWorkbook.SlicerCaches("Slicer_Payment_Method")
        .ClearManualFilter
    End With

    With ActiveWorkbook.SlicerCaches("Slicer_Buyer_Location")
        .ClearManualFilter
    End With

    With ActiveWorkbook.SlicerCaches("Slicer_Age_Group")
        .ClearManualFilter
    End With
End Sub

---
| Tool                           | Purpose                              |
| ------------------------------ | ------------------------------------ |
| **Microsoft Excel**            | Cleaning and Dashboard creation and interactivity |
| **VBA (Macros)**               | Automation of slicers and actions    |
| **Pivot Tables & Charts**      | Data aggregation and visualization   |
