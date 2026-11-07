# Business Performance Dashboard | Microsoft Excel

Interactive macro-enabled Excel dashboard for reviewing revenue, profit, quantity, sales targets and performance across products, customers and Indian regions.

## Business problem
Sales reporting from transaction-level spreadsheets makes it difficult to compare business results across time, products, regions and order statuses. This project consolidates the data into Excel PivotTables, KPI cards, charts and interactive slicers.

## Source and scope
- Workbook: `Business_Perfomance_Dashboard.xlsm`
- Sheets: `Sales Data`, `Pivot`, `Dashboard`
- Source: **8,356 records**, **20 columns**; reporting years **2025, 2026**
- Includes 14 PivotTable definitions, 16 chart definitions, slicer metadata and embedded VBA.
- Currency: INR (₹), as shown in the workbook.

## Source-wide measures (no slicers)
| Metric | Value |
|---|---:|
| Sales records | 8,356 |
| Distinct orders | 8,356 |
| Distinct customers | 18 |
| Units | 35,293 |
| Revenue | ₹127,449,475.74 |
| Cost | ₹68,427,301.43 |
| Profit | ₹22,721,520.94 |
| Profit margin | 17.83% |
| Target revenue | ₹119,890,286.67 |
| Target achievement | 106.31% |

## Data dictionary
| Field | Meaning |
|---|---|
| Order ID | Identifier for an order; check uniqueness and whether multiple lines share an ID. |
| Order Date | Date the order was recorded; Excel date serial in the source. |
| Year | Reporting year derived/stored alongside Order Date. |
| Month | Abbreviated month label; sort using a month number, not alphabetically. |
| Customer ID | Customer identifier; count distinct IDs for customer totals. |
| Customer Name | Customer label; review publication permissions. |
| Product ID | Product identifier. |
| Product Name | Product description. |
| Category | Product grouping. |
| Region | Geographic sales region. |
| State | State/territory associated with the transaction. |
| Sales Person | Sales representative label. |
| Order Status | Order fulfilment or processing state; affects realized-sales definitions. |
| Qty | Number of units in the transaction. |
| Unit Price | Recorded selling price per unit. |
| Revenue | Reported transaction revenue; validate against Qty × Unit Price. |
| Cost | Reported transaction cost; check allocation/definition. |
| Profit | Reported revenue less cost; check arithmetic consistency. |
| Target Revenue | Target/benchmark associated with the row; clarify whether targets are additive. |
| Margin % | Row-level profit / revenue; aggregate margin should be weighted by revenue. |


## Calculation examples
```excel
=[@Qty]*[@[Unit Price]]
=[@Revenue]-[@Cost]
=IFERROR([@Profit]/[@Revenue],0)
=SUM(tblSales[Revenue])
=SUM(tblSales[Profit])/SUM(tblSales[Revenue])
=ROWS(UNIQUE(tblSales[Order ID]))
=ROWS(UNIQUE(tblSales[Customer ID]))
```
The `UNIQUE` examples require a supported Excel version. PivotTable Data Model distinct counts are an alternative.

## Dashboard construction
1. Validate and format the `Sales Data` sheet; convert it into table `tblSales`.
2. Build monthly revenue, profit, quantity and target PivotTables, plus category, product, region, state, customer and order-status pivots.
3. Create KPI cards using PivotTable values and `GETPIVOTDATA`.
4. Insert PivotCharts (line, bar, doughnut or other suitable types) and arrange them on `Dashboard`.
5. Insert slicers and use **Report Connections** to connect each slicer to compatible PivotTables.
6. Optionally add the documented VBA examples for fullscreen and refresh.
7. Validate all slicer combinations and compare unfiltered totals with the source.

See `Business_Performance_Rebuild_Guide.pdf` for detailed numbered instructions and complete example VBA.

## Important limitations
- The existing Pivot KPI labeled **Total Orders** references a **Qty** pivot; the **Total Customers** KPI references a pivot order count. Review/relabel these before publishing.
- The workbook contains year 2026 data; confirm whether it is simulated, forecast or actual.
- Some source rows have inconsistent revenue or profit arithmetic; do not silently overwrite the data.
- Row-level target values may not represent independently additive targets; confirm the target definition.
- Customer and salesperson names and source distribution rights should be checked before public sharing.
- The VBA code in the guide is a reproducible example, not a verified extraction of the workbook's original macro procedures.

## Repository layout
```text
Business-Performance-Excel-Dashboard/
├── README.md
├── Business_Performance_Project_Documentation.pdf
├── Business_Performance_Rebuild_Guide.pdf
├── Business_Perfomance_Dashboard.xlsm   # only if permitted to share
└── images/dashboard.png                 # screenshot to add
```

## Tools
Microsoft Excel, PivotTables, PivotCharts, GETPIVOTDATA, slicers, formulas and optional VBA.
