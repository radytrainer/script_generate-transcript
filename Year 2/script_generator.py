import openpyxl
from docxtpl import DocxTemplate

# Load data from excel file
path = "Template Transcript Year 2 - 2024.xlsx"
workbook = openpyxl.load_workbook(path)
sheet = workbook.active
list_values = list(sheet.values)
# print(list_values)
transcript = DocxTemplate("TEMPLATE TRANSCRIPT 2024 YEAR 2.docx")

for value in list_values[1:]:
   transcript.render({
    "student_id": value[0],
    "first_name": value[1],
    "last_name": value[2],
    "sd": value[3],
    "sd_g": value[4],
    "js": value[5],
    "js_g": value[6],
    "php": value[7],
    "ph_g": value[8],
    "db": value[9],
    "db_g": value[10],
    "vc1": value[11],
    "v1_g": value[12],
    "node": value[13],
    "no_g": value[14],
    "oop1": value[15],
    "op1_g": value[16],
    "e3": value[17],
    "al_sc": value[18],
    "e3_g": value[19],
    "p3": value[20],
    "p3_g": value[21],
    "esd3": value[22],
    "es3_g": value[23],
    "t3": value[24],
    "oop2": value[25],
    "op2_g": value[26],
    "lar": value[27],
    "lar_g": value[28],
    "vue": value[29],
    "vu_g": value[30],
    "vc2": value[31],
    "v2_g": value[32],
    "e4": value[33],
    "e4_g": value[34],
    "p4": value[35],
    "p4_g": value[36],
    "esd4": value[37],
    "es4_g": value[38],
    "int": value[39],
    "int_g": value[40],
    "t5": value[41],

   })
   # print(value)
   doc_name =  str(value[1]) + "-" + str(value[2]) + "-Year 2" + ".docx"
   transcript.save(doc_name) 
