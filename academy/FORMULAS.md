# Excel formula handbook

Import sales.csv at A1 in a sheet named Sales; import products.csv at A1 in a sheet named Products. Ensure dates and numeric columns E:G are correctly recognized. Add headers Revenue, Classification, Category and Month in H1:K1. Fill H2:K2 down to row 11.

| Cell | Purpose | Formula |
|---|---|---|
| H2 | Net revenue | `=E2*F2*(1-G2)` |
| I2 | Classification | `=IF(H2>=40000,"High","Standard")` |
| J2 | Lookup | `=XLOOKUP(D2,Products!$A$2:$A$4,Products!$B$2:$B$4,"Unknown")` |
| K2 | Month | `=TEXT(B2,"mmm")` |
| M2 | Total revenue | `=SUM(H2:H11)` |
| M3 | Average revenue | `=AVERAGE(H2:H11)` |
| M4 | West revenue | `=SUMIF(C2:C11,"West",H2:H11)` |
| M5 | Laptop units | `=SUMIF(D2:D11,"Laptop",E2:E11)` |
| M6 | High orders | `=COUNTIF(I2:I11,"High")` |
| M7 | West monitor revenue | `=SUMIFS(H2:H11,C2:C11,"West",D2:D11,"Monitor")` |

Older Excel: replace XLOOKUP with `=IFERROR(VLOOKUP(D2,Products!$A$2:$B$4,2,FALSE),"Unknown")`.

## Dashboard exercise
Select A1:K11, insert a PivotTable, set Region in Rows, Product in Columns, and Sum of Revenue in Values. Insert a column chart and a Region slicer where supported. Cross-check grand totals against formulas. Format revenue as currency and discount as percentage. Validate discounts between zero and one, inspect duplicates by OrderID, and preserve original data.
