# while syntex

count = 0
while count < 10:
    count= count + 1
    print('count = %d'% count)
    if count == 10:
        print('count is 10!')

#count = 1
#count = 2
# ...
#count is 10!

prompt = """
. . . 1. Add
. . . 2. Del
. . . 3. List
. . . 4. Quit
. . .
. . . Enter number: """

number = 0
while number != 4:
    print(prompt)
    number = int(input())


bean = 10
money = 300
while money:
    print("give a bean")
    bean=bean-1
    print("bean = %d" % bean)
    if not bean:
        print("bean is zero")
        break

    
