from docxtpl import DocxTemplate, RichText, InlineImage
from docx.shared import Inches, Mm
from datetime import datetime
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, Color, colors
from copy import copy

doc = DocxTemplate('MOP_DenHaag_V1.0.docx')
context = { 'site_name' : "TestUE1" ,}
context['site_name'] = "Site1"
context['Hub_PE'] = "Site1-CR102"
context['New_Hub_PE'] = 'Site1-CR102'
doc.render(context)

doc.save("MOP_Site1_V1.1.docx")