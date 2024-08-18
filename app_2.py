import pandas as pd
from docx import Document
import pypandoc

# Load the Excel sheet
excel_file = r'/home/om/Desktop/auto maker/info_1.xlsx'  
df = pd.read_excel(excel_file)



for index, row in df.iterrows():
    print(df.iloc[index])
    doc = Document(r'/home/om/Desktop/auto maker/autofill_temp.docx')

    new_doc = doc
    for paragraph in new_doc.paragraphs:
        if '{Name}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Name}', str(row['PERSONAL INFORMATION']))
        if '{STID}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{STID}', str(row['STUDENT ID']))
        if '{EMAIL}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{EMAIL}', str(row['EMAIL ADDRESS']))
        if '{PHNO}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{PHNO}', str(row['PHONE NUMBER']))
        if '{FOS}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{FOS}', str(row['MAJOR/FEILD OF STUDY']))
        if '{INT_JOI}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{INT_JOI}', str(row['Interest and Skills']))
        if '{YOS}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{YOS}', str(row['YEAR OF STUDY']))
        if '{EXP}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{EXP}', str(row['Do you have any prior experience with robotics or related fields? (If yes, please elaborate)']))
        if '{AREA}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{AREA}', str(row['What areas of robotics are you most interested in? (You can select more']).replace("●", ""))
        if '{ANS}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{ANS}', str(row['Additional Information:']) if row['Additional Information:'] else "")
    

    docx_filename = f'/home/om/Desktop/auto maker/generated/filled_form_{str(row['PERSONAL INFORMATION'])}.docx'
    new_doc.save(docx_filename)

    pdf_filename = f'/home/om/Desktop/auto maker/member_pdf/{str(row["PERSONAL INFORMATION"])}_application.pdf'
    pypandoc.convert_file(docx_filename, 'pdf', outputfile=pdf_filename)
    
print("Forms have been successfully auto-filled.")  
