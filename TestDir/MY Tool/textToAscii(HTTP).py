# Text to ASCII (HTTP)
while True:
	print("""
	. . . Text to ASCII (HTTP)
	. . . 
	. . . Input text :""", end=" ")
	inputData=input()
	htasc=[]
	for i in range(len(inputData)):
		temp=hex(ord(inputData[i]))
		htasc.append(temp[2:])

	while len(htasc) > 0:
		char = htasc.pop(0)
		print("%"+char,end="")

