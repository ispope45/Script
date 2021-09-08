import openpyxl
from openpyxl_image_loader import SheetImageLoader

import PIL
from PIL import Image
START_LINE = 0
END_LINE = 0

if __name__ == "__main__":
    filename = "C:\\Users\\Jungly\\Downloads\\물품보유현황_20210906_1.xlsx"
    saveDir = "C:\\Users\\Jungly\\Downloads\\TEST\\"
    load_wb = openpyxl.load_workbook(filename)
    load_ws = load_wb.active

    # load_ws._images = []

    if START_LINE == 0:
        startRow = 2
    else:
        startRow = START_LINE

    if END_LINE == 0:
        totalRows = load_ws.max_row + 1
    else:
        totalRows = END_LINE

    dic = {}
    image_load = SheetImageLoader(load_ws)
    # image = image_load.get('O3')
    for i in range(startRow, totalRows):
        row = str(i)
        if image_load.image_in(f'O{row}'):
            key = load_ws[f'A{row}'].value
            val = image_load.get(f'O{row}')
            dic[key] = val

    print(dic)

    for i in range(startRow, totalRows):
        row = str(i)
        if load_ws[f'A{row}'].value in dic:
            img = openpyxl.drawing.image.Image(dic[load_ws[f'B{row}'].value])
            img.anchor = f'U{row}'
            img.width = 65
            img.height = 65
            load_ws.add_image(img)

    load_wb.save("C:\\Users\\Jungly\\Downloads\\물품보유현황_20210906_8.xlsx")
    # img = openpyxl.drawing.image.Image(image)
    # img.anchor = 'P3'
    # load_ws.add_image(img)
    # load_wb.save("C:\\Users\\Jungly\\Downloads\\물품보유현황_20210906_3.xlsx")
    #
    # image.save(saveDir + "test.png")
    # print(image)
    # print(type(image))

    # image = Image.open()
    # images = []
    # cnt = int
    # for img in load_ws._images:
    #     cnt = 1
    #     anchor = img.anchor._from
    #     data = img._data
    #     print(anchor)
    #     image = data
    #     print(help(image))
    #     image.write(saveDir + str(cnt) + ".jpg")
    #     print(type(image))
    #     image.save(saveDir + str(cnt) + ".jpg")
    #     cnt += 1

    # load_wb.save("C:\\Users\\Jungly\\Downloads\\물품보유현황_20210906_2.xlsx")