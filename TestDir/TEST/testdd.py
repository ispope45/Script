ip = '10.10.10.10/21'
subnet = int(ip.split('/')[1])
ipOctet_A = int(ip.split('/')[0].split('.')[0])
ipOctet_B = int(ip.split('/')[0].split('.')[1])
ipOctet_C = int(ip.split('/')[0].split('.')[2])
ipOctet_D = int(ip.split('/')[0].split('.')[3])

print(subnet)
print(ipOctet_A)
print(ipOctet_B)
print(ipOctet_C)
print(ipOctet_D)

ipOctet_Cp = ipOctet_C + (2 ** (24-subnet)) - 1
ipOctet_Dp = 254
gateway = f'{str(ipOctet_A)}.{str(ipOctet_A)}.{str(ipOctet_Cp)}.{str(ipOctet_Dp)}/{str(subnet)}'
rip1 = f'{str(ipOctet_A)}.{str(ipOctet_A)}.{str(ipOctet_Cp)}.{str(ipOctet_Dp-1)}/{str(subnet)}'
rip2 = f'{str(ipOctet_A)}.{str(ipOctet_A)}.{str(ipOctet_Cp)}.{str(ipOctet_Dp-2)}/{str(subnet)}'
print(gateway)
print(ipOctet_Cp)
print(rip1)
print(rip2)


aa = []

ipa = [1, 2]
ipb = ipa + [3, 4]
print(ipb)
aa.append(ipa)
aa.append(ipb)
print(aa)


for i in range(1, 2):
    print(i)


del_cnt = 3
ip_set = ['10.10.10.10/24']

cmd = list()
cmd.append('interface eth1\n')
for delCnt in range(1, del_cnt + 1):
    cmd.append('no ip add\n')

for ip in ip_set:
    cmd.append(f'ip add {ip}')

cmd.append('exit\n')
cmd.append('y\n')

print(cmd)
mask = 18
subnetmask = ''

if mask // 8 > 0:
    subnetmask += '255.'

if mask // 8 > 1:
    subnetmask += '255.'

if mask // 8 > 2:
    subnetmask += '255.'

subnetmask += str(256 - 2 ** (8 - (mask % 8))) + '.'
subnetmask += '0'

print(subnetmask)