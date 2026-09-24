from b6 import *
wh=wb.create_sheet("Help")
wh.sheet_properties.tabColor=NAVY
wh.column_dimensions["A"].width=110
wh["A1"]="DIPSEVA कसे वापरावे — 5 मिनिट मार्गदर्शक"
wh["A1"].font=Font(name="Noto Sans Devanagari",size=16,bold=True,color="FFFFFF")
wh["A1"].fill=PatternFill("solid",fgColor=NAVY); wh["A1"].alignment=CE
guide=["","1) Items मध्ये वस्तू + भाडे + Opening Stock लिहा.","2) Customers मध्ये ग्राहक + फोन + Opening Due लिहा.","3) रोजची नोंद फक्त Ledger मध्ये: Date|Customer|Type|Item|Qty|Paid.","   Issue=sामान दिले (Bill auto), Return=परत (फक्त Qty), Payment=पैसे (फक्त Paid).","4) Yogesh: Issue 850+Paid 600, Issue 900+Paid 0 => Baki 1150. Payment 500 => 650.","5) Dashboard/Khata/Stock auto. Statement मध्ये ग्राहक निवडा + WhatsApp copy.","6) App मधूनच entry करा. Sheet फक्त backup. निळ्या cell मध्ये हाताने लिहू नका."]
for i,line in enumerate(guide,start=2):
    x=wh.cell(row=i,column=1,value=line)
    x.font=BF; x.alignment=LF
    wh.row_dimensions[i].height=26 if line else 10
for sh in wb.worksheets:
    sh.sheet_properties.pageSetUpPr.fitToPage=True
print("b7 ok")
