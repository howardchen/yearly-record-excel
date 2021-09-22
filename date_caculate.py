# 篩選日期:
#   a 禮拜天不上班,去除禮拜天的日期
#   b 禮拜六必定是更換WA磁帶先挑選出來
#   c 剩餘的日期另外排序處理


import calendar

def getMothDate(year, month):
    caculated_date = []
    precaculate_date = []
    dict_tape = {0:"磁帶 A01", 1:"磁帶 A02", 2:"磁帶 A03"}
    for i in range(calendar.monthrange(year, month)[1] + 1)[1:]:
        if calendar.weekday(year, month, i) != 6:
            if calendar.weekday(year, month, i) == 5:
                caculated_date.append([i, "磁帶 WA"])
            else:
                precaculate_date.append(i)
    for i in range(len(precaculate_date)):
        caculated_date.append([precaculate_date[i], dict_tape[i%3]])
    caculated_date = sorted(caculated_date)
    for d in caculated_date:
        d[0] = str(month)+"/"+str(d[0])
    return caculated_date


# year = 2021
# month = 7
# date_list = getMothDate(year, month)
# print(date_list)