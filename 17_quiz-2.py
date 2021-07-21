from openpyxl import load_workbook
wb = load_workbook("성적목록.xlsx")
ws = wb.active

ws["H1"] = "총점"
ws["I1"] = "성적"
for col in ws.iter_cols(min_row=2, max_row= 11, min_col=4, max_col=4):
    for cell in col:
        cell.value = 10

for col in ws.iter_cols(min_row=2, max_row= 11, min_col=8, max_col=9):
    for row in col:
        row.value = "=sum(B%d:G%d)"



wb.save("scores.xlsx")
