import math
def prime(n):
    if n < 2 :
        return False
    for i in range(2,int(math.sqrt(n))+1):
        if n % i == 0 :
            return False
    return True
for _ in range(int(input())):
    x=input()
    if prime(int(x[:3])) and prime(int(x[-3:])):
        print('YES')
    else:
        print('NO')