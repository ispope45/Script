#list
a = []
a = list() # [] equal list()
b = [1,2,3]
print(b) # [1,2,3]
print(b[0]) # 1
print(b[-1]) # 3
print(b[0] + b[1]) # 1 + 2

c = [4,5,6]
print(b+c) # [1,2,3,4,5,6]
print(b*2) # [1,2,3,1,2,3]

c[2] = 0
print(c) # [4,5,0]

c[1:2] = ['a','b','c']
print(c) # [4, a, b, c, 0]

c[1:4] = []
print(c) # [4,0]

b.append(4)
print(b) # [1,2,3,4]

b.append([1,2])
print(b) # [1,2,3,4,[1,2]]

d = [4,2,1,3,5]

d.sort()
print(d) # [1,2,3,4,5]

e = ['ad','dd','bz','az']
e.sort()
print(e) # ['ad', 'az', 'bz', 'dd']

d.reverse()
print(d) # [5,4,3,2,1]

print(d.index(4)) # 1

d.insert(1,6)
print(d) # [5,6,4,3,2,1]

d.remove(6)
print(d) # [5,4,3,2,1]

print(d.pop()) # 1
print(d) # [5,4,3,2]

print(d.pop(2)) # 3
print(d) # [5,4,2]

print(d.count(5)) # 1
d.extend([1,3])
print(d) # [5,4,2,1,3]


