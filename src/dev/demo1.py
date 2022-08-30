for i in range(0,10):
    print(i)
    for j in range(0,15):
        print("j={}".format(j))
        if j== 14:
            break
        else:
            print("continue")
            continue
    if j==14:
        break
    else:
        continue
print("end")
