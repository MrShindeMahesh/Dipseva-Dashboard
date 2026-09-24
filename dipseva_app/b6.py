from b5 import *
wstm=wb.create_sheet("Statement")
wstm.sheet_properties.tabColor=GRN; wstm.freeze_panes="A5"
for col,w in zip("ABCDEFGH",[14,24,12,20,9,12,12,26]):
    wstm.column_dimensions[col].width=w
wstm.merge_cells("A1:H1")
wstm["A1"]="STATEMENT • ग्राहक बिल • वर ग्राहक निवडा"
wstm["A1"].font=Font(name="Noto Sans Devanagari",size=12,bold=True,color="FFFFFF")
wstm["A1"].fill=PatternFill("solid",fgColor=NAVY); wstm["A1"].alignment=CE
wstm.row_dimensions[1].height=28
wstm["A2"]="Select Customer:"; wstm["A2"].font=BF
wstm["B2"]="Yogesh Caterers"
wstm["B2"].font=Font(name="Noto Sans Devanagari",size=13,bold=True,color=NAVY)
wstm["B2"].fill=PatternFill("solid",fgColor="FEF3C7"); wstm["B2"].border=BD
wstm.merge_cells("B2:D2")
dd2=DataValidation(type="list",formula1="=Customers!$A$3:$A$200",allow_blank=True)
wstm.add_data_validation(dd2); dd2.add("B2")
wstm["F3"]="Total Bill"; wstm["F2"]='=SUMIFS(Ledger!$H:$H,Ledger!$B:$B,$B$2)+IFERROR(XLOOKUP($B$2,Customers!$A:$A,Customers!$F:$F),0)'
wstm["G3"]="Received"; wstm["G2"]='=SUMIFS(Ledger!$I:$I,Ledger!$B:$B,$B$2)'
wstm["H3"]="Pending"; wstm["H2"]="=$F$2-$G$2"
for cc in ["F2","G2","H2"]:
    wstm[cc].font=Font(name="Noto Sans Devanagari",size=15,bold=True,color=NAVY)
    wstm[cc].number_format=MF; wstm[cc].border=BD; wstm[cc].alignment=CE
wstm["F2"].fill=PatternFill("solid",fgColor="DBEAFE")
wstm["G2"].fill=PatternFill("solid",fgColor="D1FAE5")
wstm["H2"].fill=PatternFill("solid",fgColor="FEE2E2")
wstm["A4"]="WhatsApp:"; wstm["A4"].font=BF
wstm.merge_cells("B4:H4")
wstm["B4"]='="Namaskar "&$B$2&", DipSeva hisob: Bill Rs"&TEXT($F$2,"#,##0")&", Jama Rs"&TEXT($G$2,"#,##0")&", Baki Rs"&TEXT($H$2,"#,##0")&" - Dhanyavad"'
wstm["B4"].font=BF; wstm["B4"].fill=PatternFill("solid",fgColor="D1FAE5"); wstm["B4"].alignment=LF
wstm.row_dimensions[4].height=30
for i,h in enumerate(["Date","Type","Item ID","Item","Qty","Bill","Paid","Note"],start=1):
    x=wstm.cell(row=5,column=i,value=h)
    x.font=HF; x.fill=HFL; x.alignment=CE; x.border=BD
for r in range(6,106):
    k=r-5
    base=f'AGGREGATE(15,6,ROW(Ledger!$B$3:$B$1002)/(Ledger!$B$3:$B$1002=$B$2),{k})'
    wstm.cell(row=r,column=1).value=f'=IFERROR(INDEX(Ledger!$A:$A,{base}),"")'
    wstm.cell(row=r,column=1).number_format=DF
    wstm.cell(row=r,column=2).value=f'=IF($A{r}="","",INDEX(Ledger!$C:$C,{base}))'
    wstm.cell(row=r,column=3).value=f'=IF($A{r}="","",INDEX(Ledger!$D:$D,{base}))'
    wstm.cell(row=r,column=4).value=f'=IF($A{r}="","",INDEX(Ledger!$E:$E,{base}))'
    wstm.cell(row=r,column=5).value=f'=IF($A{r}="","",INDEX(Ledger!$F:$F,{base}))'
    wstm.cell(row=r,column=6).value=f'=IF($A{r}="","",INDEX(Ledger!$H:$H,{base}))'
    wstm.cell(row=r,column=7).value=f'=IF($A{r}="","",INDEX(Ledger!$I:$I,{base}))'
    wstm.cell(row=r,column=8).value=f'=IF($A{r}="","",INDEX(Ledger!$K:$K,{base}))'
    for col in range(1,9):
        x=wstm.cell(row=r,column=col)
        x.font=BF; x.border=BD; x.alignment=CE
        if col in (6,7): x.number_format=MF
print("b6 ok")
