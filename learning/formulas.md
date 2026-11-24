# Excel formulas and exercises

## Calculated columns
In table Sales, add Revenue: =[@Units]*[@UnitPrice]
Add Cost: =[@Units]*[@UnitCost]
Add Profit: =[@Revenue]-[@Cost]
Add Margin: =IFERROR([@Profit]/[@Revenue],0) and format as percentage.

## Summary cells
- Total revenue: =SUM(Sales[Revenue])
- Total profit: =SUM(Sales[Profit])
- Overall margin: =IFERROR(SUM(Sales[Profit])/SUM(Sales[Revenue]),0)
- West revenue: =SUMIFS(Sales[Revenue],Sales[Region],"West")
- Laptop transactions: =COUNTIFS(Sales[Product],"Laptop")
- Average transaction revenue: =AVERAGE(Sales[Revenue])
- High-value sale label: =IF([@Revenue]>=50000,"High","Standard")

## Lookup practice
Create a separate Products table with unique Product and Category columns. In modern Excel, use =XLOOKUP([@Product],Products[Product],Products[Category],"Not found"). For older versions use INDEX/MATCH with exact matching.

## Data quality
Use TRIM to remove ordinary excess spaces, inspect blanks, check duplicate rows before removing them, and verify dates after import. Do not treat missing numeric values as zero without justification.

## Expected checks
The sample has 10 transactions. Revenue is 787,500 and profit is 164,500. Verify with calculated columns and independently reconcile totals.
