#import os
def convert_date(name):
    name = name[-12:-4]
    if name[2] == '-':
        name = '20' + name
    else:
        name = name[0:4] + '-' + name[4: 6] + '-' + name[6:8]
    return name

def sort_files(files):
    dt= convert_date(files[0])
    sorted1 = [0]
    sorted2 = [dt]
    sort_len = 1
    total_len = len(files)

    while sort_len < total_len:
        dt = convert_date(files[sort_len])
        beg = 0
        end = sort_len
        pivot = sort_len // 2
        #print(pivot)
        while beg < end:
            if dt > sorted2[pivot]:
                beg = pivot + 1
            elif dt < sorted2[pivot]:
                end = pivot
            else:
                print("Error: Two files with the same date in the current directory")
                exit(1)
            pivot = (beg + end) // 2
        if pivot == 0:
            sorted1 = [sort_len] + sorted1
            sorted2 = [dt] + sorted2
        elif pivot == sort_len:
            sorted1 = sorted1 + [sort_len]
            sorted2 = sorted2 + [dt]
        else:
            sorted1 = sorted1[0: pivot] + [sort_len] + sorted1[pivot : sort_len]
            sorted2 = sorted2[0: pivot] + [dt] + sorted2[pivot : sort_len]
        sort_len += 1
        #print(sorted2)
    return [sorted1, sorted2]

#files = os.listdir('安心1号')
#print(sort_files(files))