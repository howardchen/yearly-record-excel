# 篩選日期:
#   a 禮拜天不上班,去除禮拜天的日期
#   b 禮拜六必定是更換WA磁帶先挑選出來
#   c 剩餘的日期另外排序處理

import calendar

def getMothDate(year, month):
    # 回傳日期與對應磁帶的list [['10/1', '磁帶 A01'], ['10/2', '磁帶 WA'], ...]
    holiday_date = []   # 當月份周六 周日的日期單獨處理
    working_date = [] # 除了周六 周日的日期處理
    dict_tape = {0:"磁帶 A01", 1:"磁帶 A02", 2:"磁帶 A03"}
    for i in range(calendar.monthrange(year, month)[1] + 1)[1:]:
        if calendar.weekday(year, month, i) != 6: # 禮拜天
            if calendar.weekday(year, month, i) == 5: # 禮拜六
                holiday_date.append([i, "磁帶 WA"])
            else:
                working_date.append(i) # 週六週日除外的日期就加到平常上班日
    final_date_form = [] # 最後合併所有日期的list
    final_date_form += holiday_date # 加入假日日期
    for day in range(len(working_date)): # 加入平日日期與磁帶編號
        final_date_form.append([working_date[day], dict_tape[day%3]])
    final_date_form = sorted(final_date_form) # 依照日期排序
    # 日期轉換成mm/dd格式
    for d in final_date_form:
        d[0] = str(month)+"/"+str(d[0])
    return final_date_form