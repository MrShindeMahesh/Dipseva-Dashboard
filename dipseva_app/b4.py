from b3c import *
wk=wb.create_sheet("Khata")
wk.sheet_properties.tabColor=SAF; wk.freeze_panes="A3"
wk.merge_cells("A1:G1")
wk["A1"]="KHATA • ग्राहक खाते • बेरीज (Yogesh: 850+900 बिल, 600 जमा = 1150 बाकी)"
wk["A1"].font=Font(name="Noto Sans Devanagari",size=11,bold=True,color="FFFFFF")
wk["A1"].fill=PatternFill("solid",fgColor=NAVY); wk["A1"].alignment=CE
wk.row_dimensions[1].height=30
KH=[("Customer\nग्राहक",24),("Billed ₹\nएकूण बिल",13),("Received ₹\nएकूण जमा",13),("Opening Due\nजुनी बाकी",13),("Pending ₹\nयेणे बाकी",14),("Given Qty\nदिलेले",12),("Outside Qty\nबाहेर",12)]
for i,(h,w) in enumerate(KH,start=1):
    x=wk.cell(row=2,column=i,value=h)
    x.font=HF; x.fill=HFL; x.alignment=CE; x.border=BD
    wk.column_dimensions[get_column_letter(i)].width=w
wk.row_dimensions[2].height=38
for r in range(3,203):
    wk.cell(row=r,column=1,value=f'=IF(Customers!A{r}="","",Customers!A{r})').font=BF
    wk.cell(row=r,column=2,value=f'=IF($A{r}="","",SUMIFS({LB},{LC},$A{r}))').number_format=MF
    wk.cell(row=r,column=3,value=f'=IF($A{r}="","",SUMIFS({LP},{LC},$A{r}))').number_format=MF
    wk.cell(row=r,column=4,value=f'=IF($A{r}="","",IFERROR(XLOOKUP($A{r},Customers!$A:$A,Customers!$F:$F),0))').number_format=MF
    wk.cell(row=r,column=5,value=f'=IF($A{r}="","",$B{r}+$D{r}-$C{r})').number_format=MF
    wk.cell(row=r,column=6,value=f'=IF($A{r}="","",SUMIFS({LQ},{LC},$A{r},{LT},"Issue"))').number_format='#,##0'
    wk.cell(row=r,column=7,value=f'=IF($A{r}="","",SUMIFS({LQ},{LC},$A{r},{LT},"Issue")-SUMIFS({LQ},{LC},$A{r},{LT},"Return"))').number_format='#,##0'
    for col in range(1,8):
        x=wk.cell(row=r,column=col)
        x.border=BD; x.alignment=CE
        if col==1: x.alignment=LF
wk.conditional_formatting.add("E3:E202",CellIsRule(operator="greaterThan",formula=["0"],fill=PatternFill("solid",fgColor="FEE2E2"),font=Font(color=RED,bold=True)))
wk.conditional_formatting.add("G3:G202",CellIsRule(operator="greaterThan",formula=["0"],fill=PatternFill("solid",fgColor="FEF3C7")))
wk.auto_filter.ref="A2:G202"
print("b4 ok")
