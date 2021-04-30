def findMinAndMax(L):
    max1 = float('-inf')
    min1 = float('inf')
    if L == NULL:
        return None, None
    elif len(L)==1:
        return L[0], L[0]
    else:
        for flag in L:
            if flag<min1:
                min1=flag
            if flag>max1:
                max1=flag
    return min1, max1


if findMinAndMax([]) != (None, None):
    print('测试失败!')
elif findMinAndMax([7]) != (7, 7):
    print('测试失败!')
elif findMinAndMax([7, 1]) != (1, 7):
    print('测试失败!')
elif findMinAndMax([7, 1, 3, 9, 5]) != (1, 9):
    print('测试失败!')
else:
    print('测试成功!')
