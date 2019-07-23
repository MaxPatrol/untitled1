list1 = ["a","b","c","d","e","f","g","j"]
list2 =[]
print(list1)

for i in range(10):
    a = list(c+str(i+1) for c in list1)
    list2 = list2 + a
print(list2)
