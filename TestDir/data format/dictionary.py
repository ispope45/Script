#dictionary
dic = {'name':'yang','phone':'98614520','birth':'1207'}
print(dic['name']) # yang
print(dic['birth']) # 1207

dic['age'] = '25'
print(dic) # 

del dic['age']
print(dic)

##### key is unique

print(dic.keys()) # dict_keys(['name','phone','birth'])

# for k in dic.keys() :
#    print(k)
    # name
    # phone
    # birth

print(dic.items()) # dict_items(['name','yang'....])

print(dic.get('name')) # yang

# dic['notExist_key'] keyError
print(dic.get('notExist_key')) # none

print('name' in dic) # True
print('notExist_key' in dic) # False

#value couple clear
dic.clear()
print(dic) # {}
