import cv2
from openpyxl import Workbook, load_workbook

wb = load_workbook('names.xlsx')
ws = wb.active


for i in ws.rows:
    name = i[0].value
    template = cv2.imread('template.jpg')
    if name != 'Names':
        name = name.upper()
        (width, height), baseline = cv2.getTextSize(
            name, cv2.FONT_HERSHEY_COMPLEX, 4, 4)
        cv2.putText(template, name, (1000 - width//2, 700),
                    cv2.FONT_HERSHEY_SIMPLEX, 4, (0, 0, 0), 4, cv2.LINE_AA)
        cv2.imwrite(
            f"D:\github\python automation\certificate automation\generated\{name}.jpg", template)
