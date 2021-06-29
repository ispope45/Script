# Project Euler #2
num1,num2,num3 = 1,0,0
res=0
while num3 < 4000000:
    if num3%2 == 0:
        res=res+num3
        
    num2=num1
    num1=num3
    num3=num1+num2
   
print(res)
