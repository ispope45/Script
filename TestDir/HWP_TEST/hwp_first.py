import olefile
import zlib
import struct
import openpyxl
import glob
import os
import re

CUR_PATH = os.getcwd()

def get_hwp_text(filename):
    f = olefile.OleFileIO(filename)
    dirs = f.listdir()

    # HWP 파일 검증
    if ["FileHeader"] not in dirs or \
            ["\x05HwpSummaryInformation"] not in dirs:
        raise Exception("Not Valid HWP.")

    # 문서 포맷 압축 여부 확인
    header = f.openstream("FileHeader")
    header_data = header.read()
    is_compressed = (header_data[36] & 1) == 1

    # Body Sections 불러오기
    nums = []
    for d in dirs:
        if d[0] == "BodyText":
            nums.append(int(d[1][len("Section"):]))
    sections = ["BodyText/Section" + str(x) for x in sorted(nums)]

    # 전체 text 추출
    text = ""
    for section in sections:
        bodytext = f.openstream(section)
        data = bodytext.read()
        if is_compressed:
            unpacked_data = zlib.decompress(data, -15)
        else:
            unpacked_data = data

        # 각 Section 내 text 추출
        section_text = ""
        i = 0
        size = len(unpacked_data)
        while i < size:
            header = struct.unpack_from("<I", unpacked_data, i)[0]
            rec_type = header & 0x3ff
            rec_len = (header >> 20) & 0xfff

            if rec_type in [67]:
                rec_data = unpacked_data[i + 4:i + 4 + rec_len]
                section_text += rec_data.decode('utf-16')
                section_text += "\n"

            i += 4 + rec_len

        text += section_text
        text += "\n"

    return text


if __name__ == "__main__":
    result_wb = openpyxl.Workbook()
    result_ws = result_wb.active

    SRC_PATH = CUR_PATH + "\\src"

    fileList = glob.glob(f'{SRC_PATH}\\*')
    print(fileList)
    row = 0

    hostname_reg = re.compile('\w')
    schName_reg = re.compile('학교')
    place_reg = re.compile('[가-힣]')

    for filePath in fileList:
        text = get_hwp_text(filePath)
        fileName = filePath.split("\\")[6].replace(".hwp", "")

        arr = text.split("\n")
        print(arr)
        arr_cnt = 0
        for val in arr:
            if "무선AP 설치/검수 확인서" in val:
                row += 1
                result_ws[f"A{row}"] = fileName
                print(val.split(" ")[0])
                if schName_reg.match(val.split(" ")[0]):
                    result_ws[f"B{row}"] = val.split(" ")[0]

                print(arr[arr_cnt + 8])
                if hostname_reg.match(arr[arr_cnt + 8]):
                    result_ws[f"C{row}"] = arr[arr_cnt + 8]

                print(arr[arr_cnt + 16])
                if place_reg.match(arr[arr_cnt + 16]):
                    result_ws[f"D{row}"] = arr[arr_cnt + 16]

            arr_cnt += 1

        result_wb.save(f"{CUR_PATH}\\result.xlsx")

