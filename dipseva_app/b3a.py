from b2 import *
from datetime import date
wd=wb.create_sheet("Dashboard")
wd.sheet_properties.tabColor=GRN
for col,w in zip("ABCDEFGH",[18,20,20,20,22,22,22,26]):
    wd.column_dimensions[col].width=w
wd.merge_cells("A1:H1")
wd["A1"]="DIPSEVA • डॅशबोर्ड • आजची स्थिती एका नजरेत"
wd["A1"].font=Font(name="Noto Sans Devanagari",size=16,bold=True,color="FFFFFF")
wd["A1"].fill=PatternFill("solid",fgColor=NAVY); wd["A1"].alignment=CE
wd.row_dimensions[1].height=34
wd["A2"]="सर्व आकडे Ledger मधून auto — येथे टाइप करू नका. फक्त H2 तारीख बदला."
wd["A2"].font=Font(name="Noto Sans Devanagari",size=10,color="64748B")
wd.merge_cells("A2:F2")
wd["G2"]="Today / आज:"; wd["G2"].font=BF; wd["G2"].alignment=CE
wd["H2"]=date(2026,7,15); wd["H2"].number_format=DF
wd["H2"].font=Font(name="Noto Sans Devanagari",size=12,bold=True,color=NAVY)
wd["H2"].fill=PatternFill("solid",fgColor="FEF3C7"); wd["H2"].border=BD; wd["H2"].alignment=CE
LR="Ledger!$A$3:$A$1002"; LC="Ledger!$B$3:$B$1002"; LT="Ledger!$C$3:$C$1002"
LI="Ledger!$D$3:$D$1002"; LQ="Ledger!$F$3:$F$1002"; LB="Ledger!$H$3:$H$1002"; LP="Ledger!$I$3:$I$1002"
def card(r,c,label,formula,fmt='#,##0',bg="FFFFFF"):
    wd.cell(row=r,column=c,value=label).font=Font(name="Noto Sans Devanagari",size=10,bold=True,color="64748B")
    b=wd.cell(row=r+1,column=c,value=formula)
    b.font=Font(name="Noto Sans Devanagari",size=18,bold=True,color=NAVY)
    b.fill=PatternFill("solid",fgColor=bg); b.border=BD; b.alignment=CE; b.number_format=fmt
card(3,1,"Total Billed • एकूण बिल",f"=SUM({LB})",bg="DBEAFE")
card(3,2,"Total Received • जमा",f"=SUM({LP})",bg="D1FAE5")
card(3,3,"Pending • येणे बाकी","=B4-C4",bg="FEF3C7")
card(3,5,"Today Collection • आजची जमा",f"=SUMIFS({LP},{LR},$H$2)",bg="D1FAE5")
card(3,6,"Today Billed • आजचे बिल",f"=SUMIFS({LB},{LR},$H$2)",bg="DBEAFE")
card(3,7,"Due Customers • थकबाकीदार",'=SUMPRODUCT((Khata!$E$3:$E$202>0)*(Khata!$A$3:$A$202<>""))',bg="FEE2E2")
wd["H3"]="Issue rows"; wd["H4"]=f'=COUNTIF({LT},"Issue")'
wd["H4"].font=Font(name="Noto Sans Devanagari",size=18,bold=True,color=NAVY); wd["H4"].alignment=CE
print("b3a ok")
