#for syntex
marks = [70, 40, 20, 100, 90]

number = 0
for mark in marks:
    number = number + 1
    if mark >= 60:
        print("%d student is pass" % number)
    else:
        print("%d student is failure" % number)


#9*9
for i in range(2,10):
    for j in range(1,10):
        print(i*j, end=" ")
    print('')


    
a = [(1,2), (3,4), (5,6)]
for (b,c) in a:     #b=a1 c=a2
    print(b+c)


l1 = [1,2,3,4]
res = []

for num in l1:
    res.append(num*3)

print(res)

res = [x*y for x in range(2,10) for y in range(1,10)]

print(res)
