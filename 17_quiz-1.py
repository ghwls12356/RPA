from openpyxl import Workbook
wb = Workbook()
ws = wb.active

ws.append(["학번", "출석", "퀴즈1", "퀴즈2", "중간고사", "기말고사", "프로젝트"])
for i in range(1, 2):
    ws.append([i,10,8,5,14,26,12])
    ws.append([i+1,7,3,7,15,24,18])
    ws.append([i+2,9,5,8,8,12,4])
    ws.append([i+3,7,8,7,17,21,18])
    ws.append([i+4,7,8,7,16,25,15])
    ws.append([i+5,3,5,8,8,17,0])
    ws.append([i+6,4,9,10,16,27,18])
    ws.append([i+7,6,6,6,15,19,17])
    ws.append([i+8,10,10,9,19,30,19])
    ws.append([i+9,9,8,8,20,25,20])

wb.save("성적목록.xlsx")