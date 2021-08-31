import math

A = [1,2,3,4,5,6,7,8,9,10]

print(A[1])


def lot(num1, arr):
    b = arr
    b.remove(num1)
    return b


if __name__ == "__main__":
    print(A)
    sub = lot(2, A)
    print(sub)

