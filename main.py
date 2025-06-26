import pandas as pd
from docx import Document
from docx.shared import Inches

DataFrame = pd.read_excel("./data.xls")
df = DataFrame.values
columns = DataFrame.columns

rooms = ["LIVING ROOM", "BEDROOM", "KITCHEN", "STORAGE"]

for j in range(len(df)):
    doc = Document()
    doc.add_picture("./logo.jpg", width=Inches(1.25))
    doc.add_paragraph("www.trinitycontents.com")
    doc.add_heading("FIRST INSPECTION REPORT", level=1)
    doc.add_picture("./images/home.jpg", width=Inches(2.5))
    for i in range(len(columns)):
        doc.add_heading(f"{columns[i]}", level=2)
        doc.add_paragraph(f"{df[j][i]}")
    doc.add_heading("PHOTOGRAPHS", level=2)
    for k in range(len(rooms)):
        doc.add_heading(rooms[k],level=3)
        for l in range(4):
            doc.add_picture(f"./images/{rooms[k]}/{l + 1}.jpg")
    doc.save(f"new{j}.docx")