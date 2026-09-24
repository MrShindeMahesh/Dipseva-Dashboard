from b3b import *
wd.merge_cells("A19:H19")
wd["A19"]="LOW STOCK • कमी माल"
wd["A19"].font=Font(name="Noto Sans Devanagari",size=11,bold=True,color="FFFFFF")
wd["A19"].fill=PatternFill("solid",fgColor=RED); wd["A19"].alignment=CE
wd.row_dimensions[19].height=24
for i,h in enumerate(["Item ID","Item","Opening","Given","Returned","In Hand","Min","Status"],start=1):
    x=wd.cell(row=20,column=i,value=h)
    x.font=HF; x.fill=HFL; x.alignment=CE; x.border=BD
for k,r in enumerate(range(21,29),start=1):
    wd.cell(row=r,column=1).value=f'=IFERROR(INDEX(Stock!$A:$A,{k}+2),"")'
    wd.cell(row=r,column=2).value=f'=IF($A{r}="","",XLOOKUP($A{r},Items!$A:$A,Items!$B:$B))'
    for cix,scol in zip([3,4,5,6,7],["B","C","D","E","F"]):
        wd.cell(row=r,column=cix).value=f'=IF($A{r}="","",XLOOKUP($A{r},Stock!$A:$A,Stock!${scol}:${scol}))'
    wd.cell(row=r,column=8).value=f'=IF($A{r}="","",XLOOKUP($A{r},Stock!$A:$A,Stock!$H:$H))'
    for cix in range(1,9):
        wd.cell(row=r,column=cix).border=BD
        wd.cell(row=r,column=cix).font=BF
        wd.cell(row=r,column=cix).alignment=CE
wd.merge_cells("A30:H30")
wd["A30"]="LAST 10 DAYS • रोजची जमा"
wd["A30"].font=Font(name="Noto Sans Devanagari",size=11,bold=True,color="FFFFFF")
wd["A30"].fill=PatternFill("solid",fgColor=NAVY); wd["A30"].alignment=CE
wd.row_dimensions[30].height=24
for i,h in enumerate(["Date","Billed","Received","","Customer","Billed","Received",""],start=1):
    x=wd.cell(row=31,column=i,value=h)
    x.font=HF; x.fill=PatternFill("solid",fgColor=TEAL); x.alignment=CE; x.border=BD
for k in range(10):
    r=32+k
    wd.cell(row=r,column=1).value=f"=$H$2-9+{k}"
    wd.cell(row=r,column=1).number_format="DD-MMM"
    wd.cell(row=r,column=2).value=f"=SUMIFS({LB},{LR},$A{r})"
    wd.cell(row=r,column=3).value=f"=SUMIFS({LP},{LR},$A{r})"
    wd.cell(row=r,column=5).value=f"=IFERROR(INDEX(Khata!$A:$A,{k}+3),\"\")"
    wd.cell(row=r,column=6).value=f'=IF($E{r}="","",SUMIFS({LB},{LC},$E{r}))'
    wd.cell(row=r,column=7).value=f'=IF($E{r}="","",SUMIFS({LP},{LC},$E{r}))'
    for cix in range(1,8):
        wd.cell(row=r,column=cix).border=BD
        wd.cell(row=r,column=cix).font=BF
        wd.cell(row=r,column=cix).alignment=CE
        if cix in (2,3,6,7): wd.cell(row=r,column=cix).number_format=MF
print("b3c ok")
