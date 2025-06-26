import pandas as pd
from docx import Document
from docx.shared import Inches
DataFrame = pd.read_excel("./data.xls")
df = DataFrame.values
columns = DataFrame.columns

for j in range(len(df)):
    doc = Document()
    doc.add_picture("./logo.jpg", width=Inches(1.25))
    doc.add_paragraph("www.trinitycontents.com")
    doc.add_heading("FIRST INSPECTION REPORT", level=1)
    for i in range(len(columns)):
        doc.add_heading(f"{columns[i]}", level=2)
        doc.add_paragraph(f"{df[j][i]}")
    doc.save(f"new{j}.docx")