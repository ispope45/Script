
svcChkList = [['TCP_7547', 'tcp 1-65535 7547-7547'], ['TCP_7548', 'tcp 1-65535 7547-7547']]

svcChkVal = True
for val in svcChkList:
    if svc[1] in val:
        svcChkVal = False
        svc[0] = val[0]

if svcChkVal:
    svcChkList.append(svc)


print(svcChkList)
print(svc, '\n')
