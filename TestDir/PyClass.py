# Python class
class Service:
    secret = "It is secret!"

pey = Service()
print(pey.secret)

class Service:
    secret = "Is is secret!!!"
    def sum(self, a, b):
        result =  a + b
        print("%s + %s = %s !!" % ( a, b, result))


client = Service()
client.sum(1,2)
Service.sum(client,2,3)

class Service:
    secret = "It is secret!!!!!"
    def sum(self, name, a, b):
        self.name = name
        result = a + b
        print("%s! %s + %s = %s !!!"% ( self.name, a, b, result ))

client2 = Service()
client2.sum("client2",3,5)


class Service:
    def __init__(self,name):
        self.name = name
    def sum(self,a,b):
        result = a+b
        print("%s %s + %s = %s" % ( self.name, a, b, result))

babo = Service("babo")
babo.sum(10,20)
