import win32com.client as win32

word = win32.Dispatch("Word.Application")
word.Visible = 0

# 이 주소는 절대 주소로 죽 써야 한다. 그냥 폴더내에 파일을 위치시켜도 안 먹힘.
doc1 = word.Documents.Open('D:\\Desktop\\0.doc')

# print doc1.Content.Text
# 왜 아래와 같이 변환시키는지는 디버깅해보면 쉽게 알수 있다.

test11 = str(doc1.Content)
test11 = test11.replace('Test', '서울성서초등학교')
word.ActiveDocument.SaveAs('D:\\Desktop\\1.doc')
print(test11)
word.Quit()
