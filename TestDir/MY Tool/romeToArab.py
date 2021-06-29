# rome number to arbia number

# numOfOne = {'I':1,'II':2,'III':3,'IV':4,'V':5,'VI':6,'VII':7,'VIII':8,'IX':9}
# numOfTen = {'X':10,'XX':20,'XXX':30,'XL':40,'L':50,'LX':60,'LXX':70,'LXXX':80,'XC':90}
# numOfHund = {'C':100}

inputText = input()
textList = inputText.split(' ')
#print(inputText)
print(textList)
for i in range(0,len(textList)):
    textLength = len(textList[i])
#    print(textLength)
    numSum = 0
    k=0
    for j in range(0, textLength):
        if (j+k) < textLength:
            if textList[i][j+k] == 'C':
#                if j+k+1 == textLength:
#                    break            
#                elif textList[i][j+k+1] == 'M':
#                    numSum = numSum + 900
#                    k = k + 1
#                elif textList[i][j+k+1] == 'D':
#                    numSum = numSum + 400
#                    k = k + 1
#                else:
                numSum = numSum + 100           
            elif textList[i][j+k] == 'L':
                numSum = numSum + 50     
            elif textList[i][j+k] == 'X':
                if j+k+1 == textLength:
                    break
                elif textList[i][j+k+1] == 'C':
                    numSum = numSum + 90
                    k = k + 1
                elif textList[i][j+k+1] == 'L':
                    numSum = numSum + 40
                    k = k + 1
                else:
                    numSum = numSum + 10        
            elif textList[i][j+k] == 'I':    
                if j+k+1 == textLength:
                    break
                elif textList[i][j+k+1] == 'X':
                    numSum = numSum + 9
                    k = k + 1
                elif textList[i][j+k+1] == 'V':
                    numSum = numSum + 4
                    k = k + 1
                else:
                    numSum = numSum + 1        
            elif textList[i][j+k] == 'V':
                numSum = numSum + 5
    #print("")
    print(numSum, end=" ")
    
