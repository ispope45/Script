from openpyxl import Workbook
from openpyxl.drawing.image import Image

wb = Workbook()
ws = wb.active

img = Image("C:\\Users\\Jungly\\Pictures\\중평초PoE01.JPG")
# Pillow img width : openpyxl cell width = 200 : 25 / 8 : 1
# Pillow img height : openpyxl cell height = 200 : 150 / 4 : 3

img.width = 300
img.height = 200

ws.column_dimensions['A'].width = 37.5
ws.row_dimensions[1].height = 150

ws.add_image(img, "A1")

wb.save(filename="C:\\Users\\Jungly\\Pictures\\test4.xlsx")
