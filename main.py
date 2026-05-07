import csv
from openpyxl import Workbook

datos = []
with open("datos.txt", "r", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        datos.append(row)

wb = Workbook()
ws = wb.active
ws.title = "RFC Data"

ws.append(["RFC", "RAZON SOCIAL", "CODIGO POSTAL"])

for d in datos:
    ws.append([d["RFC"], d["RAZON_SOCIAL"], d["CODIGO_POSTAL"]])

wb.save("resultado.xlsx")
print("Excel generado: resultado.xlsx")