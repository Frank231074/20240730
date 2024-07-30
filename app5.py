import openpyxl
from openpyxl.styles import Font

wb = openpyxl.Workbook()

ws = wb.active

font = Font(name="Arial", size=24, bold=True, italic=True, color="FF0000")

cell = ws["A1"]
cell.value = "Hello world"
cell.font = font

wb.save("files/font.xlsx")
