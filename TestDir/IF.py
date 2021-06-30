# if syntex
money = 1
if money:
    print('i have money')
else:
    print('i haven\'t money')


if money:
    print('1')
#print('2')   syntex error
    print('3')

if money:
    print('1')
    print('2')
#        print('3') syntex error

print(1 < 2) # True123
print(1 > 2) # False
print(1 == 2) # False
print(1 != 2) # True
print(1 <= 2) # True
print(1 >= 2) # False

print(1 or 0) # True
print(1 and 0) # False
print(not 1) # False

string = 'Life is short'
print('L' in string) # True
print('L' not in string) # False

l1 = [1,2,3]
print(1 in l1) # True

if 0 in l1 :
    print('0 in l1')
elif 1 in l1 :
    print('1 in l1')
else :
    print('nope')


if 1 in l1 :
    pass
else:
    print("else!")

    
