import codecs


list_eng = {'a', 'A', 'b', 'B', 'c', 'C', 'd', 'D', 'e', 'E', 'f', 'F', 'g', 'G', 'h', 'H', 'i', 'I', 'j', 'J', 'k',
            'K', 'l', 'L', 'm', 'M', 'n',
            'N', 'o', 'O', 'p', 'P', 'q', 'Q', 'r', 'R', 's', 'S', 't', 'T', 'u',
            'U', 'v', 'V', 'w', 'W', 'x', 'X', 'y', 'Y', 'z', 'Z', '0', '1', '2', '3', '4',
            '5', '6', '7', '8', '9', '{', '}', '_', '-'}
list_full = []
fileObj = codecs.open("D://CTF_task.txt", "r", encoding="utf-8")
len_str = 0
for line in fileObj:
    list_answer = []
    for sym in line:
        if (sym in list_eng) == True:
            list_answer.append(sym)

    list_full.append(list_answer)

len_str = 0
for elem in list_full:
    if len(elem) > len_str:
        len_str = len(elem)
        print("".join(elem))

fileObj.close()