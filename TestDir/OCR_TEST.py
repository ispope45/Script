from PIL import Image
import pytesseract
import cv2


def ocr(imgFile, lang='eng'):
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    print(imgFile)
    img = cv2.imread(imgFile, 0)
    cv2.imshow('image', img)
    print(type(img))

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    print(type(gray))
    # img = Image.open(imgFile)
    text = pytesseract.image_to_string(gray, lang=lang)

    print('+++ OCR Result +++')
    print(text)
    return text


text = ocr(r'C:\Users\Jungly\Desktop\양진이\7_(가칭)신길중학교_PoE01_전면.jpg', lang='kor+eng')
print(type(text))
