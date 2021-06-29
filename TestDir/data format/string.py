#String format
first = "Life"
second = " is"
third = " short"
print(first + second + third)   # Life is short
print(first + second + third*2) # Life is short short
print("\n") # new line
print("\v") # vertical tab
print("\t") # horison tab
print("\r") # carrige return
print("\f") # form feed
print("\a") # alarm
print("\b") # back space
print("\000") # null
print("\\") # \
print("\'") # '
print("\"") # "

#example
print("=" * 50)
print("Life is short")
print("=" * 50)

#string slice
string = first + second + third # Life is short
print(string[0]) # L
print(string[1]) # i
print(string[2]) # f
print(string[3]) # e

print(string[-0]) # L
print(string[-1]) # t
print(string[-2]) # r
print(string[-3]) # o

print(string[0:4]) # Life
print(string[0:]) # Life is short
print(string[8:]) # short
print(string[:]) # Life is short

#string formatting
print("I`m %d years." %25) # I`m 25 years.

# %s string
# %c character
# %d integer
# %f floating-point
# %o oct
# %x hex
# %% %

print("l'm {0} years." . format(25) ) # I'm 25 years
print("{0} {1} {2}" . format(10, 20 ,30)) # 10 20 30

#sorting
print("{0:>10}".format("hello")) #          hello
print("{0:<10}".format("hello")) # hello
print("{0:^10}".format("hello")) #      hello

#lower to upper
print(string.upper()) # LIFE IS SHORT
print(string)

#upper to lower
print(string.lower()) # life is short
print(string)

#count
print(string.count('i')) # 2

#location find
print(string.find("e")) # 3 : -1
print(string.index("e")) # 3 : error

#join
let = "*"
print(let.join("abcd")) # a*b*c*d

#delete space
print(string.lstrip()) # Life is short
print(string.rstrip()) # Life is short
print(string.strip()) # Life is short

#replace & split
print(string.replace("Life", "Your leg")) # your leg is short
      
