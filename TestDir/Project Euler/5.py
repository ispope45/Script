# Project Euler #5
#i=0
#isbool = True
#chkCnt=0
#while isbool:
#    i=i+1
#    for a in range(1,16):
#        if chkCnt == 15:
#            isbool=False
#        elif i%a == 0 :
#            chkCnt=chkCnt+1
#        else :
#            chkCnt=0
#            break
#
#
#print(i-1)

def primeFactorsOfFactors(n):
    l = []
    # Create a list of all factors
    div = list(range(2,n+1))

    while len(div) > 0:
        # x takes the first value of the factors and remove it
        try:
            x = div.pop(0)
        except:
            pass
        # We add this factor in the output list
        l.append(x)
        # We check the nexts factor by comparison to the actuals factors
        i = 0
        for i in range(len(div)):
            if not div[i] % x:
                div[i] = int(div[i]/x)
    return l

# We check all prime factors for our factor
l = primeFactorsOfFactors(20)

# Let's multiply all those numbers
x = 1
for n in l:
    x *= n

print(x)
