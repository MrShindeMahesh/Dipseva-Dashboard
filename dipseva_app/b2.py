from b1 import *
wi=wb.create_sheet("Items")
wi.sheet_properties.tabColor=TEAL; wi.freeze_panes="A3"
wi.merge_cells("A1:G1")
wi["A1"]="ITEMS • वस्तू यादी • भाडे येथूनच Ledger मध्ये auto येईल"
wi["A1"].font=Font(name="Noto Sans Devanagari",size=12,bold=True,color="FFFFFF")
wi["A1"].fill=PatternFill("solid",fgColor=TEAL); wi["A1"].alignment=CE
wi.row_dimensions[1].height=26
IH=[("Item ID\nकोड",12),("Item Name\nवस्तू (EN)",24),("Item Marathi\nवस्तू (मराठी)",24),("Rate ₹\nभाडे",11),("Opening Stock\nसुरुवातीचा माल",15),("Min Stock\nकिमान",10),("Note",20)]
for i,(h,w) in enumerate(IH,start=1):
    x=wi.cell(row=2,column=i,value=h)
    x.font=HF; x.fill=HFL; x.alignment=CE; x.border=BD
    wi.column_dimensions[get_column_letter(i)].width=w
wi.row_dimensions[2].height=36
for r in range(3,203):
    for col in range(1,8):
        x=wi.cell(row=r,column=col)
        x.font=BF; x.border=BD; x.alignment=CE if col not in (2,3,7) else LF
wc=wb.create_sheet("Customers")
wc.sheet_properties.tabColor=SAF; wc.freeze_panes="A3"
wc.merge_cells("A1:F1")
wc["A1"]="CUSTOMERS • ग्राहक यादी • नाव एकदाच लिहा"
wc["A1"].font=Font(name="Noto Sans Devanagari",size=12,bold=True,color="FFFFFF")
wc["A1"].fill=PatternFill("solid",fgColor=NAVY); wc["A1"].alignment=CE
wc.row_dimensions[1].height=26
CH=[("Customer\nग्राहक",26),("Contact Person\nव्यक्ती",22),("Phone\nफोन",15),("Type\nमंडप/केटरर्स",15),("Address / Marathi Name\nपत्ता",30),("Opening Due ₹\nजुनी बाकी",14)]
for i,(h,w) in enumerate(CH,start=1):
    x=wc.cell(row=2,column=i,value=h)
    x.font=HF; x.fill=HFL; x.alignment=CE; x.border=BD
    wc.column_dimensions[get_column_letter(i)].width=w
wc.row_dimensions[2].height=36
for r in range(3,203):
    for col in range(1,7):
        x=wc.cell(row=r,column=col)
        x.font=BF; x.border=BD; x.alignment=CE if col not in (1,2,5) else LF
        if col==6: x.number_format=MF
dd=DataValidation(type="list",formula1='"Mandap,Caterers,Walk-in,Other"',allow_blank=True)
wc.add_data_validation(dd); dd.add("D3:D202")
print("b2 ok")
