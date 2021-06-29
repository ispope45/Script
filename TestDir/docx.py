import win32com.client
word = win32com.client.DispatchEx("Word.Application")
word.Visible = True
word.DisplayAlerts = 0
word.Documents.Open("D:\\Desktop\\0.docx")

word.Selection.Find.Text = "Test"
word.Selection.Find.Replacement.Text = "서울성서초등학교22"
word.Selection.Find.Forward = True
word.Selection.Find.Execute(Replace=2)

word.ActiveDocument.SaveAs('D:\\Desktop\\1.docx')
word.Quit() # releases Word object from memory