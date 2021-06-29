def transNumber(base,n):
    text = '0123456789abcdefghijklmnopqrstuvwxyz'
    nText = text[:n]
    li = list(nText)
    baseLi = list(base)
    baseLength = len(baseLi)
    i=0
    while i < baseLength:
        result = ( n * li.index(baseLi[i]) ) + li.index(baseLi[i+1])
        print(result, end=" ")
        i = i + 2
	
    print("")
    return 0
        
while True:
    inputData = input("Input Number : ")
    inputNum = input("Transform : ")
    transNum = int(inputNum)
    transNumber(inputData, transNum)
    
