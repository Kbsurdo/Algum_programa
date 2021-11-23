import random

data = ['', '', '', '', '', '', '', '', '', '']
for aux1 in range(6, 2002, 2):    
    for aux2 in range(10):
        data[aux2] = random.randint(1, 1000)
    print(data)
