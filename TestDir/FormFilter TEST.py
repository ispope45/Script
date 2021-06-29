import re
def form_filter(formData):
    repList = ["<", " ", "=", "\\", "/", "..", "+", "%", "*", "#", "--", "?", ">"]
    print(type(formData))
    for elem in repList:
        print(type(elem))
        formData = formData.replace(elem, "")
        print(formData)
    return formData

formData= " flg ks le    4   0 9 2 ;d\ 0fg ,. >>    <  <  ..f<> fg;dfg"
form_filter(formData)