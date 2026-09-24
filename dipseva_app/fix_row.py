import openpyxl
p = "/home/shaarx/Documents/Dipseva/Dipseva_Ledger_NEW.xlsx"
wb = openpyxl.load_workbook(p)
ws = wb["Ledger"]
# restore template formulas on the cleaned test row + verify contiguous block
for r in range(3, 1003):
    a = ws.cell(r, 1).value; b = ws.cell(r, 2).value; c = ws.cell(r, 3).value
    if a in (None, "") and b in (None, "") and c in (None, ""):
        ws.cell(r, 5).value = f'=IF($D{r}="","",IFERROR(XLOOKUP($D{r},Items!$A:$A,Items!$B:$B),"ERR"))'
        ws.cell(r, 7).value = f'=IF($D{r}="","",IFERROR(XLOOKUP($D{r},Items!$A:$A,Items!$D:$D),0))'
        ws.cell(r, 8).value = f'=IF($C{r}="Issue",IFERROR($F{r}*$G{r},0),0)'
# count data rows
n = sum(1 for r in range(3, 1003) if not (ws.cell(r, 1).value in (None, "") and ws.cell(r, 2).value in (None, "") and ws.cell(r, 3).value in (None, "")))
print("data rows:", n)
wb.save(p)
print("formulas restored")
