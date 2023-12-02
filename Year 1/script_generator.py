import openpyxl
from docxtpl import DocxTemplate

# Load data from excel file
path = "Template Transcript Year 1 - 2024.xlsx"
workbook = openpyxl.load_workbook(path)
sheet = workbook.active
list_values = list(sheet.values)
# print(list_values)
transcript = DocxTemplate("TRANSCRIPT 2024 YEAR 1.docx")

for value in list_values[1:]:
   transcript.render({
    "student_id": value[0],
    "first_name": value[1],
    "last_name": value[2],
    "logic": value[3],
    "l_g": value[4],
    "bcum": value[5],
    "bc_g": value[6],
    "design": value[7],
    "d_g": value[8],
    "p1": value[9],
    "p1_g": value[10],
    "e1": value[11],
    "e1_g": value[12],
    "esd1": value[13],
    "es1_g": value[14],
    "t1": value[15],
    "wd": value[16],
    "wd_g": value[17],
    "algo": value[18],
    "al_g": value[19],
    "p2": value[20],
    "p2_g": value[21],
    "e2": value[22],
    "e2_g": value[23],
    "esd2": value[24],
    "es2_g": value[25],
    "t2": value[26]

   })
   doc_name =  str(value[1]) + "-" + str(value[2])+ "-Year 1" + ".docx"
   transcript.save(doc_name) 