# Project Euler #4
cnt=0
res=0
for n1 in range(100,1000):
    for n2 in range(100,1000):
        chk=str(n1*n2)
        chkCnt=len(chk)
        for n3 in range(0,chkCnt):
            if chk[n3] == chk[chkCnt-(n3+1)]:
                cnt=cnt+1

        if cnt == chkCnt:
            temp=int(chk)
            if temp > res :
                res=temp
                num1,num2=n1,n2
        cnt=0
        
print(num1,num2,res)
