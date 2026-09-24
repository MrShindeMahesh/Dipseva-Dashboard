from b3a import *
wd.merge_cells("A6:H6")
wd["A6"]="TOP PENDING CUSTOMERS • जास्त बाकी असलेले ग्राहक"
wd["A6"].font=Font(name="Noto Sans Devanagari",size=11,bold=True,color="FFFFFF")
wd["A6"].fill=PatternFill("solid",fgColor=TEAL); wd["A6"].alignment=CE
wd.row_dimensions[6].height=24
tops=["#","Customer / ग्राहक","Billed","Received","Pending","Given","Returned","Outside"]
for i,h in enumerate(tops,start=1):
    x=wd.cell(row=7,column=i,value=h)
    x.font=HF; x.fill=HFL; x.alignment=CE; x.border=BD
for r in range(8,18):
    wd.cell(row=r,column=1,value=f"={r-7}").font=BF
    for cix in range(1,9):
        wd.cell(row=r,column=cix).border=BD
        wd.cell(row=r,column=cix).font=BF
        wd.cell(row=r,column=cix).alignment=CE
    wd.cell(row=r,column=2).value=f'=IFERROR(INDEX(Khata!$A:$A,MATCH(LARGE(Khata!$E$3:$E$202,{r-7}),Khata!$E$3:$E$202,0)),"")'
    wd.cell(row=r,column=3).value=f'=IF($B{r}="","",SUMIFS({LB},{LC},$B{r})+XLOOKUP($B{r},Customers!$A:$A,Customers!$F:$F,0))'
    wd.cell(row=r,column=4).value=f'=IF($B{r}="","",SUMIFS({LP},{LC},$B{r}))'
    wd.cell(row=r,column=5).value=f'=IF($B{r}="","",$C{r}-$D{r})'
    wd.cell(row=r,column=6).value=f'=IF($B{r}="","",SUMIFS({LQ},{LC},$B{r},{LT},"Issue"))'
    wd.cell(row=r,column=7).value=f'=IF($B{r}="","",SUMIFS({LQ},{LC},$B{r},{LT},"Return"))'
    wd.cell(row=r,column=8).value=f'=IF($B{r}="","",$F{r}-$G{r})'
    for cix in (3,4,5): wd.cell(row=r,column=cix).number_format=MF
wd.conditional_formatting.add("E8:E17",CellIsRule(operator="greaterThan",formula=["0"],fill=PatternFill("solid",fgColor="FEE2E2"),font=Font(color=RED,bold=True)))
print("b3b ok")
