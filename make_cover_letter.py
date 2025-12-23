import zipfile, io, base64

# --- Your Cover Letter Text ---
text = """Cover Letter — Why I Should Be Hired for This Role

I believe I am a strong fit for this Data Analytics Intern position because my academic experience, hands-on project work, and technical skill set closely align with the responsibilities of the role. I have worked extensively on collecting, cleaning, and analyzing large datasets using Python, SQL, and Excel—skills essential for supporting data-driven decisions.

Through projects such as Retail Sales Analysis, AI Diagnostic Tool (Healthcare Analytics), and E-commerce Customer Behavior Analysis, I gained practical experience performing EDA, identifying trends and anomalies, and building dashboards using Tableau and Power BI.

Additionally, my experience presenting insights, preparing documentation, and collaborating in teams strengthens my ability to communicate technical information clearly. I am detail-oriented and committed to maintaining data accuracy and integrity throughout the analysis process.

I am excited to contribute to your data projects, stay updated with analytics best practices, and support the team in developing impactful, data-driven solutions. My passion for analytics and strong problem-solving skills make me a dedicated and reliable candidate for this internship.

Thank you for considering my application.
"""

# ---- DOCX XML CREATION ----
def paragraph(text):
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"

paras = "".join([paragraph(line) for line in text.split("\n")])

document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        {paras}
    </w:body>
</w:document>
"""

content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" 
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" 
    Target="word/document.xml"/>
</Relationships>
"""

# ---- BUILD THE DOCX ----
with zipfile.ZipFile("Cover_Letter.docx", "w") as docx:
    docx.writestr("[Content_Types].xml", content_types)
    docx.writestr("_rels/.rels", rels)
    docx.writestr("word/document.xml", document_xml)

print("Your file 'Cover_Letter.docx' has been created successfully!")
