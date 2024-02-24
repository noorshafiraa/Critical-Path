'''
CRITICAL PATH FINDER BY NOOR SHAFIRA

IMPORTANT: INDEX 0 EVERY LIST = NODE 1 IN DIAGRAM
ASSUME 1 AS STARTING NODE

Input form (activity, init, final, duration) example
# Activity: A
# init: 1
# final: 1
# duration: 3

Output form: (activity, init, final, duration, ES, EF, LS, LF, TF, FF, IF, isCritical) example
# Activity: A
# init: 1
# final: 1
# duration: 3
# ES: 0
# EF: 3
# LS: 0
# LF: 3
# TF: 0
# FF: 0
# IF: 0
# isCritical: True
'''
import pandas as pd

# INITIALIZE CLASS AND LIST
class task:
    def __init__(self, activity, init, final, duration):
        self.activity = activity
        self.init = init 
        self.final = final
        self.duration = duration

class result:
    def __init__(self, ES, EF, LS, LF, TF, FF, IF, isCritical):  
        self.ES = ES
        self.EF = EF
        self.LS = LS
        self.LF = LF
        self.TF = TF
        self.FF = FF
        self.IF = IF
        self.isCritical = isCritical

list = [] #task list
list1 = [] #result list
list2 = [] #final list
list3 = [] #final list inside list
ES = []
LF = []

# DECLARATION
column = 4 #total column input file
endnode = 0

# INPUT ======================================================================
# Reading 1st sheet on excel file
df1 = pd.read_excel('data.xlsx')

# Reading each activity (row)
total = (len(df1))
for i in range(total):
    activity = df1.iloc[i, 0]
    init = int(df1.iloc[i, 1])
    final = int(df1.iloc[i, 2])
    duration = int(df1.iloc[i, 3])

    # Finding end node
    if (final >  endnode):
        endnode = final

    # Appending instances to list
    list.append(task(activity, init, final, duration))

# PROCESS ======================================================================
# Finding EET (EET = ES)

#initializing ES list with 0
for i in range(endnode):
    ES.append(0)

for i in range(endnode):
    # Checking every activity with initial node i
    for j in range(total):
        if list[j].init == i+1:
            temp = ES[i] + list[j].duration
            if temp > ES[(list[j].final)-1]:
                ES[(list[j].final)-1] = temp

# Finding LET (LET = LF)

#initializing LF list with end node EET value
for i in range(endnode):
    LF.append(ES[endnode-1])

for i in range(endnode-1, -1, -1):
    # Checking every activity with initial node i
    for j in range(total-1, -1, -1):
        if list[j].final == i+1:
            temp = LF[i] - list[j].duration
            if temp < LF[(list[j].init)-1]:
                LF[(list[j].init)-1] = temp

# Finding EF, LS, Float, isCritical
for i in range(total):
    duration = list[i].duration
    ES_val = ES[(list[i].init)-1]
    EF = ES_val + duration
    LF_val = LF[(list[i].final)-1]
    LS = LF_val - duration
    TF = LF_val - ES_val - duration
    FF = ES[(list[i].final)-1] - ES_val - duration
    IF = ES[(list[i].final)-1] - LF[(list[i].init)-1] - duration

    #Critical requirement TF = 0
    if (TF == 0):
        isCritical = True
    else:
        isCritical = False
    
    # Appending instances to list1
    list1.append(result(ES_val, EF, LS, LF_val, TF, FF, IF, isCritical))

# OUTPUT =====================================================================
# Appending each activity data to list2 temporary
for i in range(total):
    list2.append(list[i].activity)
    list2.append(list[i].init)
    list2.append(list[i].final)
    list2.append(list[i].duration)
    list2.append(list1[i].ES)
    list2.append(list1[i].EF)
    list2.append(list1[i].LS)
    list2.append(list1[i].LF)
    list2.append(list1[i].TF)
    list2.append(list1[i].FF)
    list2.append(list1[i].IF)
    list2.append(list1[i].isCritical)

    # Appending to list in list (list3)
    list3.append(list2)
    list2 = []

# Writing data to excel file result.xlsx
result_df = pd.DataFrame( data=list3, columns=['Activity', 'Initial', 'Final', 'Duration', 
                                               'ES', 'EF', 'LS', 'LF', 'TF', 'FF', 'IF', 'isCritical'])

with pd.ExcelWriter('result.xlsx') as writer: 
    result_df.to_excel(
        writer, 
        sheet_name='Sheet1', 
        index=False
    )
