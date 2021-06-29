# Sets **
# Unordered
# Not duplication
# Not indexing

s1 = set([1,2,3])
l1 = list(s1)

print(s1) # {1,2,3}
print(l1) # [1,2,3]

s2 = set([2,3,4])
print(s1 & s2) # {2,3}
print(s1.intersection(s2)) # {2,3}

print(s1 | s2) # {1,2,3,4}
print(s1.union(s2)) # {1,2,3,4}

print(s1 - s2) # {1}
print(s1.difference(s2)) # {1}

s1.add(4) # just 1 value
print(s1) # {1,2,3,4}

s1.update([5,6,7]) # mutiple values
print(s1) # {1,2,3,4,5,6,7}

s1.remove(5)
print(s1) # {1,2,3,4,6,7}

