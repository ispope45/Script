# Project Euler #1
#mul3 = int(1000 / 3)
#mul5 = int(1000 / 5)
#mul15 = int(1000 / 15)
#res=0
#
#for i in range(1,mul3+1):
#    res=res+(3 * i)
#
#for j in range(1,mul5):
#    res=res+(5 * j)
#
#for k in range(1,mul15+1):
#    res=res-(15 * k)
#print(res)

num, res1 = 1, 0
while num < 1000:
    if num%3 == 0:
        res1=res1+num
    elif num%15 == 0:
        pass
    elif num%5 == 0:
        res1=res1+num
    num=num+1
    
print(res1)
        
