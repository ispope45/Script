#function
def sum(a,b):
    return a + b

print(sum(1,2))

def say():
    return "Life is short!"

print(say())

str=say()
print(str)


def sum1(a,b):
    c = a + b
    print("a + b = %d" %c)

a = sum1(3,4) # function call 
print(a) # a is None

sum1(2,4) # function call


#input data
def calc_sum(*args):
    sum=0
    for i in args:
        sum = sum + i
    return sum
print(calc_sum(1,2,3,4,5,6,7,8,9,10))


# calculator

def calc():
    text = """
    . . . 1. sum
    . . . 2. sub
    . . . 3. mul
    . . . 4. div
    . . . Enter Accumulator : """

    text2 = """
    . . . input Number : """
    acc=0
    i=0
    j=0
    while acc == 0:
        print(text,end=" ")
        acc=int(input())
    while i == 0:
        print(text2,end=" ")
        i=int(input())
    while j == 0:
        print(text2,end=" ")
        j=int(input())
    
    if acc != 0:
        if i != 0:
            if j != 0:
                if acc == 1:
                    result = i + j
                elif acc == 2:
                    result = i - j
                elif acc == 3:
                    result = i * j
                elif acc == 4:
                    result = i / j
    print(result)
    
calc()

def double_acc(a,b):
    return a+b, a*b

test = double_acc(1,5)
print(test)
sum1 , mul1 = double_acc(5,5)
print(sum1)
print(mul1)

def TFtest(TF=True):
    if TF:
        print('True')
    else :
        print('False')


TFtest()
TFtest(False)


def recurPower(base, exp):
    if exp <=0:
        return 1
    return base * recurPower(base, exp-1)

recurPower(2,3) # 8   2**3

