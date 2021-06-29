# Project Euler #3

num = 600851475143

isbool=True
i=1
while isbool:
    i=i+1
    if num%i == 0:
        num=num/i
        print(i)
    if num < i:
        break
    

    
