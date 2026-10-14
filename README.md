# Basic Sales Summary — Level 1, Day 14

## Objective
Calculate total sales, average sales per transaction, and transaction count using Excel.

## Files
- `Basic_Sales_Summary.xlsx` — workbook with Sales Data and Summary sheets
- `retail_sales.csv` — synthetic sample data
- `summary_sheet.csv` — KPI summary and formulas
- `README.md` — instructions and verification

## Steps
1. Open the Excel workbook.
2. Review the `Sales Data` sheet. Sales_Amount = Quantity × Unit_Price.
3. Open the `Summary` sheet and review the three KPIs.
4. Use these formulas if recreating the summary:
   - Total Sales: `=SUM('Sales Data'!F2:F11)`
   - Average Sales: `=AVERAGE('Sales Data'!F2:F11)`
   - Transaction Count: `=COUNTA('Sales Data'!A2:A11)`
5. Verify by adding the 10 sales amounts: 100, 100, 700, 240, 60, 150, 100, 700, 50, 360.

## Verified results for this synthetic demo dataset
- Total Sales: ₹2,560
- Average Sales per Transaction: ₹256.00
- Transaction Count: 10

## Notes
This is synthetic demo data, not real company data. The transaction count uses non-empty Transaction_ID cells. If one transaction spans multiple rows, count unique transaction IDs instead.
