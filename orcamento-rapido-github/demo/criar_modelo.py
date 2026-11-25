"""Gera demo/modelo.xlsx: modelo de orçamento fictício (o cliente real usa o dele)."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
ws = wb.active
ws.title = "Orcamento"
ws["B1"] = "DISTRIBUIDORA EXEMPLO LTDA"
ws["B1"].font = Font(size=16, bold=True)
ws["B3"], ws["E3"] = "ORÇAMENTO", "DATA:"
for col, titulo in zip("BCDEFG", ["CÓDIGO", "DESCRIÇÃO", "QTD", "VALOR UNI.", "ESTOQUE", "TOTAL"]):
    ws[f"{col}8"] = titulo
    ws[f"{col}8"].font = Font(bold=True)
ws["F41"], ws["G41"] = "TOTAL", "=SUM(G9:G40)"
ws["F41"].font = ws["G41"].font = Font(bold=True)
for col, w in zip("BCDEFG", [16, 42, 8, 14, 10, 14]):
    ws.column_dimensions[col].width = w
wb.save(Path(__file__).with_name("modelo.xlsx"))
