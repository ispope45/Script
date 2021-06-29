a = 1
b = 2

print(a) # 1
print(a is b) # false
print(a is 1) # true

del(a)
del(b)
#print(a) # error : not define

a, b = 2, 3
print(a) # 2
print(b) # 3

(c,d) = (4,5)
[e,f] = [6,7]
print(a,b,c,d,e,f) # 2, 3, 4, 5, 6, 7

# list variable
l1 = [1,2,3]
l2 = l1
l1[1] = 4
print(l1 is l2) # True

print(l1) # 1,4,3
print(l2) # 1,4,3

# l1 < l2
# l1` < l2

l3 = l1[:]
l1[1] = 2
print(l3 is l1) # False

print(l1) # 1,2,3
print(l3) # 1,4,3

from copy import copy
l4 = copy(l1)
print(l1 is l4) # False

l1[1] = 4
print(l4) # 1,2,3
print(l1) # 1,4,3



