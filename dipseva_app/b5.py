from b4 import *
wst=wb.create_sheet("Stock")
wst.sheet_properties.tabColor=TEAL; wst.freeze_panes="A3"
wst.merge_cells("A1:H1")
wst["A1"]="STOCK • माल हिशोब • In Hand = Opening + Return - Given"
wst["A1"].font=Font(name="Noto Sans Devanagari",size=11,bold=True,color="FFFFFF")
wst["A1"].fill=PatternFill("solid",fgColor=NAVY); wst["A1"].alignment=CE
wst.row_dimensions[1].height=30
SH=[("Item ID",11),("Opening",10),("Given\nदिले",10),("Returned\nपरत",10),("In Hand\nहातात",10),("Min\nकिमान",9),("Revenue ₹\nकमाई",13),("Status",16)]
for i,(h,w) in enumerate(SH,start=1):
    x=wst.cell(row=2,column=i,value=h)
    x.font=HF; x.fill=HFL; x.alignment=CE; x.border=BD
    wst.column_dimensions[get_column_letter(i)].width=w
for r in range(3,203):
    wst.cell(row=r,column=1,value=f'=IF(Items!A{r}="","",Items!A{r})')
    wst.cell(row=r,column=2,value=f'=IF($A{r}="","",IFERROR(XLOOKUP($A{r},Items!$A:$A,Items!$E:$E),0))')
    wst.cell(row=r,column=3,value=f'=IF($A{r}="","",SUMIFS({LQ},{LI},$A{r},{LT},"Issue"))')
    wst.cell(row=r,column=4,value=f'=IF($A{r}="","",SUMIFS({LQ},{LI},$A{r},{LT},"Return"))')
    wst.cell(row=r,column=5,value=f'=IF($A{r}="","",$B{r}+$D{r}-$C{r})')
    wst.cell(row=r,column=6,value=f'=IF($A{r}="","",IFERROR(XLOOKUP($A{r},Items!$A:$A,Items!$F:$F),0))')
    wst.cell(row=r,column=7,value=f'=IF($A{r}="","",SUMIFS({LB},{LI},$A{r}))')
    wst.cell(row=r,column=8,value=f'=IF($A{r}="","",IF($E{r}<$F{r},"LOW",IF($E{r}=0,"EMPTY","OK")))')
    for col in range(1,9):
        x=wst.cell(row=r,column=col)
        x.font=BF; x.border=BD; x.alignment=CE
        if col==7: x.number_format=MF
wst.conditional_formatting.add("E3:E202",FormulaRule(formula=['AND($A3<>"",$E3<$F3)'],fill=PatternFill("solid",fgColor="FEE2E2")))
wst.auto_filter.ref="A2:H202"
print("b5 ok")
