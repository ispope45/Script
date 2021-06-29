# file in/out

f=open("E:/Python/test.txt",'w')    # file open
                                    # r = read, w = write, a = addition
f.close() # file close

f=open("E:/Python/newFile.txt",'w')
for i in range(1,11):
    data = "%d Line \n" % i
    f.write(data)

f.close()

f=open("E:/Python/newFile.txt",'r')
line = f.readline()
print(line)
f.close()

f=open("E:/Python/newFile.txt",'r')
while True:
    line = f.readline()
    if not line: break
    print(line, end=" ")

f.close()

f=open("E:/Python/newFile.txt",'r')
lines = f.readlines()
for line in lines:
    print(line , end=" ")

f.close()

f=open("E:/Python/newFile.txt",'r')
data=f.read()
print(data)
f.close()

# context add in file
f=open("E:/Python/newFile.txt",'a')
for i in range(11,20):
    data = "%d Line \n" %i
    f.write(data)

f.close()

with open("E:/Python/newFile2.txt",'w') as f:   # file open
    f.write("Life is too short, you need python!")
    #file close

with open("E:/Python/newFIle3.py",'w') as f:
    f.write("import sys\n")
    f.write("args = sys.argv[1:]\n")
    f.write("for i in args:\n")
    f.write("\tprint(i)\n")



