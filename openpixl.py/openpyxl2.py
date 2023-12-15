import openpyxl

wb = openpyxl.load_workbook('teste.xlsx')

ws = wb.active

dados = ws.values
for linha in dados:
    print(linha)
