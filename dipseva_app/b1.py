from b0 import *
HDRS=[("Date\nतारीख",13),("Customer\nग्राहक",22),("Type\nप्रकार",12),("Item ID\nकोड",10),("Item Name\nवस्तू",20),("Qty\nनग",9),("Rate\nभाडे",10),("Bill ₹\nबिल",12),("Paid ₹\nजमा",12),("Pay Mode\nपेमेंट",14),("Note\nनोंद",24)]
ws.merge_cells("A1:K1")
c=ws["A1"]; c.value="DIPSEVA LEDGER • Issue=भाड्याने दिले • Return=परत आले • Payment=पैसे जमा"
c.font=Font(name="Noto Sans Devanagari",size=11,bold=True,color="FFFFFF")
c.fill=PatternFill("solid",fgColor=SAF); c.alignment=CE
ws.row_dimensions[1].height=30
for i,(h,w) in enumerate(HDRS,start=1):
    x=ws.cell(row=2,column=i,value=h)
    x.font=HF; x.fill=HFL; x.alignment=CE; x.border=BD
    ws.column_dimensions[get_column_letter(i)].width=w
ws.row_dimensions[2].height=38
for r in range(3,N+1):
    ws.row_dimensions[r].height=20
    for col in range(1,12):
        x=ws.cell(row=r,column=col)
        x.font=BF; x.border=BD
        x.alignment=CE if col not in (2,5,11) else LF
    ws.cell(row=r,column=5).value=f'=IF($D{r}="","",IFERROR(XLOOKUP($D{r},Items!$A:$A,Items!$B:$B),"ERR"))'
    ws.cell(row=r,column=7).value=f'=IF($D{r}="","",IFERROR(XLOOKUP($D{r},Items!$A:$A,Items!$D:$D),0))'
    ws.cell(row=r,column=8).value=f'=IF($C{r}="Issue",IFERROR($F{r}*$G{r},0),0)'
    ws.cell(row=r,column=1).number_format=DF
    ws.cell(row=r,column=6).number_format='#,##0'
    ws.cell(row=r,column=7).number_format=MF
    ws.cell(row=r,column=8).number_format=MF
    ws.cell(row=r,column=9).number_format=MF
d1=DataValidation(type="custom",formula1='OR(NOT(ISERROR(DATEVALUE(A3))),AND(ISNUMBER(A3),LEFT(CELL("format",A3))="D"))',allow_blank=True)
ws.add_data_validation(d1); d1.add(f"A3:A{N}")
d2=DataValidation(type="list",formula1="=Customers!$A$3:$A$200",allow_blank=True)
ws.add_data_validation(d2); d2.add(f"B3:B{N}")
d3=DataValidation(type="list",formula1='"Issue,Return,Payment"',allow_blank=True)
ws.add_data_validation(d3); d3.add(f"C3:C{N}")
d4=DataValidation(type="list",formula1="=Items!$A$3:$A$200",allow_blank=True)
ws.add_data_validation(d4); d4.add(f"D3:D{N}")
d5=DataValidation(type="list",formula1='"Cash,UPI,PhonePe SK,Paytm SK,PhonePe VK,PAYTM B QR,Bank"',allow_blank=True)
ws.add_data_validation(d5); d5.add(f"J3:J{N}")
ws.conditional_formatting.add(f"C3:C{N}",FormulaRule(formula=['$C3="Issue"'],fill=PatternFill("solid",fgColor="DBEAFE")))
ws.conditional_formatting.add(f"C3:C{N}",FormulaRule(formula=['$C3="Return"'],fill=PatternFill("solid",fgColor="D1FAE5")))
ws.conditional_formatting.add(f"C3:C{N}",FormulaRule(formula=['$C3="Payment"'],fill=PatternFill("solid",fgColor="FEF3C7")))
ws.auto_filter.ref=f"A2:K{N}"
print("b1 ok")
