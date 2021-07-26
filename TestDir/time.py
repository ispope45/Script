import time
import datetime

print(time.time())

print(time.localtime(time.time()))

print(datetime.datetime.today())

dt = datetime.datetime.today()
print(dt.microsecond)
print(str(dt.year) + str(dt.month) + str(dt.day) + str(dt.hour) + str(dt.minute) + str(dt.second) + str(dt.microsecond)[0:3])