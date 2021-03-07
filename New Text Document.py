from random import randint
from xlutils.copy import copy
from xlrd import open_workbook
import xlrd
import pandas as pd
import import_ipynb
import os

os.system('python -m pip install xlrd==1.2.0')

try:
    import pandas
except:
    os.system('python -m pip install pandas')
try:
    import random
except:
    os.system('python -m pip install random')
try:
    import xlwt
except:
    os.system('python -m pip install xlwt')
try:
    import import_ipynb
except:
    os.system('python -m pip install import_ipynb')
try:
    import xlutils
except:
    os.system('python -m pip install xlutils')
try:
    import shutil
except:
    os.system('python -m pip install shutil')

#!pip install random
#!pip install xlwt
#!pip install xlrd==1.2.0


def split100(regNoList, nameList, markList, outputFileName):
    import random

    def splitMarks(ans, marksPartC, marksPartB, marksPartA):
        originalMarks = ans
        ans = int(ans)

        # Part C Marks Allocation
        for i in range(1):
            partCmark = random.randint(0, 15)
            if partCmark+marksPartC[0] > 15:
                break
            if partCmark >= ans:
                ans = partCmark
                return (marksPartC, marksPartB, marksPartA)
            marksPartC[0] += partCmark
            ans -= partCmark

        # Part B Marks Allocation
        partBmark = random.randint(0, 13)
        for i in range(5):
            if partBmark+sum(marksPartB) > 65:
                break
            if(ans-partBmark < 0):
                partBmark = ans
                break
            marksPartB[i] += partBmark
            ans -= partBmark

        # Part A Marks Allocation
        for i in range(10):
            partAmark = random.randint(0, 2)
            if partAmark+sum(marksPartA) > 20:
                break
            if(ans-partAmark < 0):
                partAmark = ans
                break
            marksPartA[i] += partAmark
            ans -= partAmark
        return (marksPartA, marksPartB, marksPartC)

    def adjustMarks(temp, a, b, c):
        while True:
            for i in range(1):
                if temp == 0:
                    return (a, b, c)
                if c[i] < 16:
                    c[i] += 1
                    temp -= 1
            for j in range(5):
                if temp == 0:
                    return (a, b, c)
                if b[j] < 13:
                    b[j] += 1
                    temp -= 1
            for k in range(10):
                if temp == 0:
                    return (a, b, c)
                if a[k] < 2:
                    a[k] += 1
                    temp -= 1
        return (a, b, c)
    flag = 0
    while flag == 0:
        try:
            writeToExcel = []
            for i in markList:
                marksPartC = [0]
                marksPartB = [0 for i in range(5)]
                marksPartA = [0 for i in range(10)]

                try:
                    int(i)
                    (a, b, c) = splitMarks(i, marksPartC, marksPartB, marksPartA)
                    if sum(a+b+c) < int(i):
                        temp = int(i) - sum(a+b+c)
                        (a, b, c) = adjustMarks(temp, a, b, c)
                    writeToExcel.append(a+b+c+[int(i)])
                except:
                    marksPartC = ['AB']
                    marksPartB = ['AB' for i in range(5)]
                    marksPartA = ['AB' for i in range(10)]
                    writeToExcel.append(
                        marksPartA+marksPartB+marksPartC+['AB'])
            flag = 1
        except:
            flag = 0
    temp = []
    for i in range(len(writeToExcel)):
        single = []
        for j in range(10):
            single.append(writeToExcel[i][j])
        for j in range(10, 15):
            aORb = random.randint(0, 1)
            if str(writeToExcel[i][j]) == 'AB':
                single.append('AB')
                single.append('AB')
            elif aORb == 0:
                single.append(0)
                single.append(writeToExcel[i][j])
            else:
                single.append(writeToExcel[i][j])
                single.append(0)
        for j in range(15, len(writeToExcel[i])):
            single.append(writeToExcel[i][j])
        temp.append(single)

    writeToExcel = temp
    import xlwt
    from xlwt import Workbook

    wb = Workbook()

    sheet1 = wb.add_sheet('Sheet 1')

    for i in range(10):
        sheet1.write(0, i+2, i+1)

    counter = 12

    for i in range(11, 16):
        sheet1.write(0, counter, str(i)+str('a'))
        sheet1.write(0, counter+1, str(i)+str('b'))
        counter += 2

    sheet1.write(0, counter, 16)

    for i in range(len(temp)):
        sheet1.write(i+1, 0, nameList[i])

    for i in range(len(temp)):
        sheet1.write(i+1, 1, str(regNoList[i]))

    for i in range(len(temp)):
        for j in range(len(temp[i])):
            try:
                sheet1.write(i+1, j+2, int(temp[i][j]))
            except:
                sheet1.write(i+1, j+2, str(temp[i][j]))

    sheet1.write(0, j+2, "Total Marks")

    fileName = "With Part C 100 "+str(outputFileName)+".xls"
    wb.save(fileName)


def split100noC(regNoList, nameList, markList, outputFileName):
    import random

    def splitMarks(ans, marksPartB, marksPartA):
        originalMarks = ans
        ans = int(ans)

        # Part B Marks Allocation
        partBmark = random.randint(0, 16)
        for i in range(5):
            if partBmark+sum(marksPartB) > 65:
                break
            if(ans-partBmark < 0):
                partBmark = ans
                break
            marksPartB[i] += partBmark
            ans -= partBmark

        # Part A Marks Allocation
        for i in range(10):
            partAmark = random.randint(0, 2)
            if partAmark+sum(marksPartA) > 20:
                break
            if(ans-partAmark < 0):
                partBmark = ans
                break
            marksPartA[i] += partAmark
            ans -= partAmark
        return (marksPartA, marksPartB)

    def adjustMarks(temp, a, b,):
        while True:
            for j in range(5):
                if temp == 0:
                    return (a, b)
                if b[j] < 16:
                    b[j] += 1
                    temp -= 1
            for k in range(10):
                if temp == 0:
                    return (a, b)
                if a[k] < 2:
                    a[k] += 1
                    temp -= 1
        return (a, b)
    writeToExcel = []

    for i in markList:
        marksPartB = [0 for i in range(5)]
        marksPartA = [0 for i in range(10)]

        try:
            int(i)
            (a, b) = splitMarks(i, marksPartB, marksPartA)
            if sum(a+b) < int(i):
                temp = int(i) - sum(a+b)
                (a, b) = adjustMarks(temp, a, b)
            writeToExcel.append(a+b+[int(i)])
        except:
            marksPartB = ['AB' for i in range(5)]
            marksPartA = ['AB' for i in range(10)]
            writeToExcel.append(marksPartA+marksPartB+['AB'])

    temp = []
    for i in range(len(writeToExcel)):
        single = []
        for j in range(10):
            single.append(writeToExcel[i][j])
        for j in range(10, 15):
            aORb = random.randint(0, 1)
            if str(writeToExcel[i][j]) == 'AB':
                single.append('AB')
                single.append('AB')
            elif aORb == 0:
                single.append(0)
                single.append(writeToExcel[i][j])
            else:
                single.append(writeToExcel[i][j])
                single.append(0)
        for j in range(15, len(writeToExcel[i])):
            single.append(writeToExcel[i][j])
        temp.append(single)

    writeToExcel = temp
    import xlwt
    from xlwt import Workbook

    wb = Workbook()

    sheet1 = wb.add_sheet('Sheet 1')

    for i in range(10):
        sheet1.write(0, i+2, i+1)
    counter = 12
    for i in range(11, 16):
        sheet1.write(0, counter, str(i)+str('a'))
        sheet1.write(0, counter+1, str(i)+str('b'))
        counter += 2
    sheet1.write(0, counter, "Total Marks")

    for i in range(len(temp)):
        sheet1.write(i+1, 0, nameList[i])

    for i in range(len(temp)):
        sheet1.write(i+1, 1, str(regNoList[i]))

    for i in range(len(temp)):
        for j in range(len(temp[i])):
            try:
                sheet1.write(i+1, j+2, int(temp[i][j]))
            except:
                sheet1.write(i+1, j+2, str(temp[i][j]))

    fileName = "Without Part C 100 "+str(outputFileName)+".xls"
    wb.save(fileName)


def split60(regNoList, nameList, markList, outputFileName):
    import random

    def splitMarks(ans, marksPartC, marksPartB, marksPartA):
        originalMarks = ans
        ans = int(ans)

        # Part C marks Allocation
        for i in range(1):
            partCmark = random.randint(0, 8)
            if partCmark+marksPartC[0] > 8:
                break
            if partCmark >= ans:
                ans = partCmark
                return (marksPartA, marksPartB, marksPartC)
            marksPartC[0] += partCmark
            ans -= partCmark

        # Part B Marks Allocation
        for i in range(2):
            partBmark = random.randint(0, 16)
            if partBmark+sum(marksPartB) > 32:
                break
            if(ans-partBmark < 0):
                partBmark = ans
                break
            marksPartB[i] += partBmark
            ans -= partBmark

        # Part A Marks Allocation
        for i in range(10):
            partAmark = random.randint(0, 2)
            if partAmark+sum(marksPartA) > 10:
                break
            if(ans-partAmark < 0):
                partAmark = ans
                break
            marksPartA[i] += partAmark
            ans -= partAmark
        return (marksPartA, marksPartB, marksPartC)

    def adjustMarks(temp, a, b, c):
        while True:
            for j in range(1):
                if temp == 0:
                    return (a, b, c)
                if c[j] < 8:
                    c[j] += 1
                    temp -= 1
            for j in range(2):
                if temp == 0:
                    return (a, b, c)
                if b[j] < 16:
                    b[j] += 1
                    temp -= 1
            for k in range(10):
                if temp == 0:
                    return (a, b, c)
                if a[k] < 2:
                    a[k] += 1
                    temp -= 1
        return (a, b, c)
    writeToExcel = []

    for i in markList:
        marksPartC = [0]
        marksPartB = [0 for i in range(2)]
        marksPartA = [0 for i in range(10)]

        try:
            int(i)
            (a, b, c) = splitMarks(i, marksPartC, marksPartB, marksPartA)
            if sum(a+b+c) < int(i):
                temp = int(i) - sum(a+b+c)
                (a, b, c) = adjustMarks(temp, a, b, c)
            writeToExcel.append(a+b+c+[int(i)])
        except:
            marksPartC = ['AB']
            marksPartB = ['AB' for i in range(2)]
            marksPartA = ['AB' for i in range(10)]
            writeToExcel.append(marksPartA+marksPartB+marksPartC+['AB'])

    temp = []
    for i in range(len(writeToExcel)):
        single = []
        for j in range(10):
            single.append(writeToExcel[i][j])
        for j in range(10, 13):
            aORb = random.randint(0, 1)
            if str(writeToExcel[i][j]) == 'AB':
                single.append('AB')
                single.append('AB')
            elif aORb == 0:
                single.append(0)
                single.append(writeToExcel[i][j])
            else:
                single.append(writeToExcel[i][j])
                single.append(0)
        for j in range(13, len(writeToExcel[i])):
            single.append(writeToExcel[i][j])
        temp.append(single)

    writeToExcel = temp
    import xlwt
    from xlwt import Workbook

    wb = Workbook()

    sheet1 = wb.add_sheet('Sheet 1')

    for i in range(10):
        sheet1.write(0, i+2, i+1)
    counter = 12
    for i in range(11, 14):
        sheet1.write(0, counter, str(i)+str('a'))
        sheet1.write(0, counter+1, str(i)+str('b'))
        counter += 2
    sheet1.write(0, counter, "Total Marks")

    for i in range(len(temp)):
        sheet1.write(i+1, 0, nameList[i])

    for i in range(len(temp)):
        sheet1.write(i+1, 1, str(regNoList[i]))

    for i in range(len(temp)):
        for j in range(len(temp[i])):
            try:
                sheet1.write(i+1, j+2, int(temp[i][j]))
            except:
                sheet1.write(i+1, j+2, str(temp[i][j]))

    fileName = "Without Part C 60 "+str(outputFileName)+".xls"
    wb.save(fileName)


#!pip install import_ipynb
#!pip install pandas
#!pip install os
# from Split import *

exna = ""
for i in os.listdir():
    if i.startswith("Question Number CO Mapping"):
        exna = i


def generate_excel(excel_file_name):
    def find_splitting_type(name, t):
        splitting_type = 0

        df_CO = pd.read_excel(exna, sheet_name="Sheet1")
        df_CO.head()
        index_of_subject = 0
        for i in range(len(df_CO)):
            if str(df_CO.iloc[i, 0]) == name:
                index_of_subject = i

        counter = 0
        for i in df_CO.iloc[index_of_subject+t, :]:
            if pd.isna(i): break
            if i == 'NIL':
                splitting_type = 3
                break
            counter += 1

        if counter == 17:
            splitting_type = 3
        elif counter == 21:
            splitting_type = 2
        elif counter == 23:
            splitting_type = 1

        return splitting_type

    df_Paper = pd.read_excel(excel_file_name)
    df_Paper.head()

    column_names = []
    for i in df_Paper.columns[3:]:
        if str(i).startswith("Unnamed"):
            break
        column_names.append(i)

    nameList = []
    for i in range(len(df_Paper)):
        if not pd.isna(df_Paper.loc[i, df_Paper.columns[2]]):
            nameList.append(df_Paper.loc[i, df_Paper.columns[2]])

    regNoList = []
    for i in range(len(df_Paper)):
        if not pd.isna(df_Paper.loc[i, df_Paper.columns[1]]):
            try:
                temp = int(df_Paper.loc[i, df_Paper.columns[1]])
                regNoList.append(temp)
            except:
                regNoList.append(str(df_Paper.loc[i, df_Paper.columns[1]]))

    for col_name in column_names:
        markList = []
        splitting_type = 0
        for i in range(len(df_Paper)):
            if not pd.isna(df_Paper.loc[i, col_name]):
                if df_Paper.loc[i, col_name] != 'AB':
                    markList.append(int(df_Paper.loc[i, col_name]))
                if df_Paper.loc[i, col_name] == 'AB':
                    markList.append(df_Paper.loc[i, col_name])
        if 'FIAT' in excel_file_name:
            t = 1
            splitting_type = find_splitting_type(col_name, t)
        elif 'SIAT' in excel_file_name:
            t = 2
            splitting_type = find_splitting_type(col_name, t)
        elif 'MODEL' in excel_file_name:
            t = 3
            splitting_type = find_splitting_type(col_name, t)

        if splitting_type == 1:

            outputFileName = excel_file_name[:-5] + \
                str(' splitted ')+col_name.replace("/", " ")

            split100(regNoList, nameList, markList, outputFileName)

        elif splitting_type == 2:

            outputFileName = excel_file_name[:-5] + \
                str(' splitted ')+col_name.replace("/", " ")

            split100noC(regNoList, nameList, markList, outputFileName)

        elif splitting_type == 3:

            outputFileName = excel_file_name[:-5] + \
                str(' splitted ')+col_name.replace("/", " ")

            split60(regNoList, nameList, markList, outputFileName)


myList = []
for i in os.listdir():
    if (i.startswith("6") or i.startswith("8") or i.startswith("1") or i.startswith("2") or i.startswith("4") or i.startswith("3") or i.startswith("5") or i.startswith("7")) and "UNIV" not in i:
        myList.append(i)
for i in myList:
    generate_excel(i)
    print("\nGenerating excel for "+str(i)+" ...\n")


print("Splitting done")

#!pip install pandas
#!pip install xlrd==1.2.0
#!pip install xlutils
#!pip install import_ipynb
# import Excel_Generator


def analyze_100(excel_file_name):
    df_Paper = pd.read_excel(excel_file_name)
    df_Paper.head()

    column_names = df_Paper.columns[2:]
    total_no_students = len(df_Paper.loc[1:, 1])
    whole_data = []
    for col_name in column_names[:-1]:
        temp = []
        attended = 0
        sixty_percent_2m = 0
        sixty_percent_13m = 0
        sixty_percent_16m = 0

        for i in df_Paper.loc[1:, col_name]:
            temp_data = []
            temp.append(i)
            flag = 0
            try:
                i = int(i)
                flag = 1
            except:
                flag = 0
                pass
            if flag == 1:
                i = int(i)
                if i > 0:
                    attended += 1
                try:
                    if col_name <= 10:
                        if i > 1:
                            sixty_percent_2m += 1
                    elif col_name == 16 or col_name == '16a' or col_name == '16b':
                        if i >= 9:
                            sixty_percent_16m += 1
                except:
                    if int(str(col_name)[:2]) >= 11 and int(str(col_name)[:2]) <= 15:
                        if i > 7:
                            sixty_percent_13m += 1
                    elif col_name == 16 or col_name == '16a' or col_name == '16b':
                        if i >= 9:
                            sixty_percent_16m += 1
        try:
            if col_name <= 10:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_2m)
                temp_data.append(sixty_percent_2m/attended*100)
            else:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_16m)
                temp_data.append(sixty_percent_16m/attended*100)
        except:
            if int(str(col_name)[:2]) >= 11 and int(str(col_name)[:2]) <= 15:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_13m)
                temp_data.append(sixty_percent_13m/attended*100)
            else:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_16m)
                temp_data.append(sixty_percent_16m/attended*100)
        whole_data.append(temp_data)
    from xlrd import open_workbook
    from xlutils.copy import copy

    rb = open_workbook(excel_file_name)

    wb = copy(rb)

    s = wb.get_sheet(0)

    i = 0

    for myRow in range(total_no_students+3, total_no_students+7):
        j = 0
        for myCol in range(2, 23):
            s.write(myRow, myCol, whole_data[j][i])
            j += 1
        i += 1

    s.write(total_no_students+4, 0, "Number of students attempted")
    s.write(total_no_students+5, 0, "Number of Students Scored >= 60% of Marks")
    s.write(total_no_students+6, 0,
            "Percentage of Students Scored >= 60% of Marks")
    s.write(total_no_students+7, 0,
            "Average number of students scored >=60% marks")
    s.write(total_no_students+8, 0, "CO Attainment Level")
    wb.save(excel_file_name)


def analyze_without_c_100(excel_file_name):
    df_Paper = pd.read_excel(excel_file_name)
    df_Paper.head()

    column_names = df_Paper.columns[2:]
    total_no_students = len(df_Paper.loc[1:, 1])
    whole_data = []
    for col_name in column_names[:-1]:
        temp = []
        attended = 0
        sixty_percent_2m = 0
        sixty_percent_16m = 0

        for i in df_Paper.loc[1:, col_name]:
            temp_data = []
            temp.append(i)
            flag = 0
            try:
                i = int(i)
                flag = 1
            except:
                flag = 0
                pass
            if flag == 1:
                i = int(i)
                if i > 0:
                    attended += 1
                try:
                    if col_name <= 10:
                        if i > 1:
                            sixty_percent_2m += 1
                except:
                    if int(str(col_name)[:2]) >= 11 and int(str(col_name)[:2]) <= 16:
                        if i >= 9:
                            sixty_percent_16m += 1
        try:
            if col_name <= 10:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_2m)
                temp_data.append(sixty_percent_2m/attended*100)
            else:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_16m)
                temp_data.append(sixty_percent_16m/attended*100)
        except:
            if int(str(col_name)[:2]) >= 11 and int(str(col_name)[:2]) <= 15:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_16m)
                temp_data.append(sixty_percent_16m/attended*100)
            else:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_16m)
                temp_data.append(sixty_percent_16m/attended*100)
        whole_data.append(temp_data)

    from xlrd import open_workbook
    from xlutils.copy import copy

    rb = open_workbook(excel_file_name)

    wb = copy(rb)

    s = wb.get_sheet(0)

    i = 0

    for myRow in range(total_no_students+3, total_no_students+7):
        j = 0
        for myCol in range(2, 22):
            s.write(myRow, myCol, whole_data[j][i])
            j += 1
        i += 1

    s.write(total_no_students+4, 0, "Number of students attempted")
    s.write(total_no_students+5, 0, "Number of Students Scored >= 60% of Marks")
    s.write(total_no_students+6, 0,
            "Percentage of Students Scored >= 60% of Marks")
    s.write(total_no_students+7, 0,
            "Average number of students scored >=60% marks")
    s.write(total_no_students+8, 0, "CO Attainment Level")
    wb.save(excel_file_name)


def analyze_60(excel_file_name):
    df_Paper = pd.read_excel(excel_file_name)
    df_Paper.head()

    column_names = df_Paper.columns[2:]
    total_no_students = len(df_Paper.loc[1:, 1])
    whole_data = []
    for col_name in column_names[:-1]:
        temp = []
        attended = 0
        sixty_percent_2m = 0
        sixty_percent_13m = 0
        sixty_percent_16m = 0

        for i in df_Paper.loc[1:, col_name]:
            temp_data = []
            temp.append(i)
            flag = 0
            try:
                i = int(i)
                flag = 1
            except:
                flag = 0
                pass
            if flag == 1:
                i = int(i)
                if i > 0:
                    attended += 1
                try:
                    if col_name <= 10:
                        if i > 1:
                            sixty_percent_2m += 1
                    elif col_name == 13 or col_name == '13a' or col_name == '13b':
                        if i >= 5:
                            sixty_percent_13m += 1
                except:
                    if int(str(col_name)[:2]) >= 11 and int(str(col_name)[:2]) <= 12:
                        if i > 9:
                            sixty_percent_16m += 1
                    elif col_name == 13 or col_name == '13a' or col_name == '13b':
                        if i >= 5:
                            sixty_percent_13m += 1
        try:
            if col_name <= 10:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_2m)
                temp_data.append(sixty_percent_2m/attended*100)
            elif int(str(col_name)[:2]) >= 11 and int(str(col_name)[:2]) <= 12:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_16m)
                temp_data.append(sixty_percent_16m/attended*100)
            else:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_13m)
                temp_data.append(sixty_percent_13m/attended*100)
        except:
            if int(str(col_name)[:2]) >= 11 and int(str(col_name)[:2]) <= 12:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_16m)
                temp_data.append(sixty_percent_16m/attended*100)
            elif int(str(col_name)[:2]) >= 11 and int(str(col_name)[:2]) <= 12:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_16m)
                temp_data.append(sixty_percent_16m/attended*100)
            else:
                temp_data.append(col_name)
                temp_data.append(attended)
                temp_data.append(sixty_percent_13m)
                temp_data.append(sixty_percent_13m/attended*100)
        whole_data.append(temp_data)
    from xlrd import open_workbook
    from xlutils.copy import copy

    rb = open_workbook(excel_file_name)

    wb = copy(rb)

    s = wb.get_sheet(0)

    i = 0

    for myRow in range(total_no_students+3, total_no_students+7):
        j = 0
        for myCol in range(2, 18):
            s.write(myRow, myCol, whole_data[j][i])
            j += 1
        i += 1
    s.write(total_no_students+4, 0, "Number of students attempted")
    s.write(total_no_students+5, 0, "Number of Students Scored >= 60% of Marks")
    s.write(total_no_students+6, 0,
            "Percentage of Students Scored >= 60% of Marks")
    s.write(total_no_students+7, 0,
            "Average number of students scored >=60% marks")
    s.write(total_no_students+8, 0, "CO Attainment Level")
    wb.save(excel_file_name)


def find_splitting_type(excel_file_name):
    if "With Part C 100" in excel_file_name:
        analyze_100(excel_file_name)
    elif "Without Part C 100" in excel_file_name:
        analyze_without_c_100(excel_file_name)
    elif "Without Part C 60" in excel_file_name:
        analyze_60(excel_file_name)
    print("\nAnalyzing excel "+str(excel_file_name)+" ...\n")


myList = []
for i in os.listdir():
    if i.startswith("With"):
        myList.append(i)
for i in myList:
    find_splitting_type(i)


#!pip install pandas
#!pip install import_ipynb
#!pip install random


exna = ""
for i in os.listdir():
    if i.startswith("Question Number CO Mapping"):
        exna = i


def get_CO_mapping(subject_name):
    df_CO = pd.read_excel(exna, sheet_name="Sheet1")
    df_CO.head()
    index_of_subject = 0
    for i in range(len(df_CO)):
        if str(df_CO.iloc[i, 0]) == subject_name:
            index_of_subject = i
    whole_List = []
    for ctr in range(1, 4):
        CO_List = []
        for i in df_CO.iloc[index_of_subject+ctr, :]:
            if pd.isna(i):
                break
            CO_List.append(i)
        whole_List.append(CO_List)

    print(whole_List)
    return whole_List


def map_CO(excel_file_name):
    df_Paper = pd.read_excel(excel_file_name)
    subject_name = ""
    flag = 0
    for i in excel_file_name[-5::-1]:
        if i == " " and flag == 0:
            subject_name += ""
            flag = 1
        elif i == " " and flag == 1:
            subject_name = subject_name[::-1].replace(" ", "/")
            break
        subject_name += i

    CO_List = get_CO_mapping(subject_name)

    total_no_students = len(df_Paper.loc[1:, 1])

    from xlrd import open_workbook
    from xlutils.copy import copy

    rb = open_workbook(excel_file_name)

    wb = copy(rb)

    s = wb.get_sheet(0)

    i = 0
    if 'FIAT' in excel_file_name: i = 0
    elif 'SIAT' in excel_file_name: i = 1
    elif 'MODEL' in excel_file_name: i = 2

    for myRow in range(total_no_students+2, total_no_students+3):
        j = 0
        for myCol in range(1, len(CO_List[i])+1):
            s.write(myRow, myCol, CO_List[i][j])
            j += 1

    wb.save(excel_file_name[:-4]+".xls")


def adjust_16m(excel_file_name):
    import xlrd
    from xlrd import open_workbook
    from xlutils.copy import copy
    from random import randint

    rb = open_workbook(excel_file_name)

    sh = rb.sheet_by_index(0)

    wb = copy(rb)

    s = wb.get_sheet(0)

    length = 0
    mark = []
    total_mark = []
    for i in range(1, sh.nrows):
        if sh.cell_value(i, 22) == '':
            length = i
            break
        mark.append(sh.cell_value(i, 22))
        total_mark.append(sh.cell_value(i, 23))

    i = 0
    col = 22
    a16, b16, a1660, b1660 = 0, 0, 0, 0
    for row in range(1, length):
        aORb = randint(0, 1)
        if mark[i] == 'AB':
            s.write(row, col, 'AB')
            s.write(row, col+1, 'AB')
        elif aORb == 0:
            s.write(row, col, mark[i])
            s.write(row, col+1, 0)
            if mark[i] > 0:
                a16 += 1
            if mark[i] > 9:
                a1660 += 1
        else:
            s.write(row, col, 0)
            s.write(row, col+1, mark[i])
            if mark[i] > 0:
                b16 += 1
            if mark[i] > 9:
                b1660 += 1

        s.write(row, col+2, total_mark[i])
        i += 1

    s.write(length+1, 22, "16a")
    s.write(length+2, 22, a16)
    s.write(length+3, 22, a1660)
    s.write(length+4, 22, a1660/a16*100)
    s.write(length+1, 23, "16b")
    s.write(length+2, 23, b16)
    s.write(length+3, 23, b1660)
    s.write(length+4, 23, b1660/b16*100)
    s.write(0, 22, '16a')
    s.write(0, 23, "16b")
    s.write(0, 24, "Total Marks")

    wb.save(excel_file_name)


myList = []
for i in os.listdir():
    if i.startswith("With"):
        myList.append(i)
for i in myList:
    if "With Part C 100" in i:
        adjust_16m(i)
    map_CO(i)
    print("\nMapping CO for "+str(i)+" ...\n")


#!pip install import_ipynb
#!pip install xlrd==1.2.0
#!pip install xlutils
#!pip install random

# import CO_Mapper


assna = ""
for i in os.listdir():
    if i.startswith("Assignment"):
        assna = i


def calc_CO(excel_file_name):

    rb = open_workbook(excel_file_name)

    sh = rb.sheet_by_index(0)

    wb = copy(rb)

    s = wb.get_sheet(0)

    length = 0
    c1, c2, c3, c4, c5, c6 = [], [], [], [], [], []

    for i in range(1, sh.nrows):
        if str(sh.cell_value(i, 3)).startswith("CO"):
            length = int(i)
            break
    for i in range(2, sh.ncols):
        if str(sh.cell_value(length, i)).strip() == 'CO1':
            if sh.cell_value(length-3, i) != '':
                c1.append(sh.cell_value(length-3, i))
        elif str(sh.cell_value(length, i)).strip() == 'CO2':
            if sh.cell_value(length-3, i) != '':
                c2.append(sh.cell_value(length-3, i))
        elif str(sh.cell_value(length, i)).strip() == 'CO3':
            if sh.cell_value(length-3, i) != '':
                c3.append(sh.cell_value(length-3, i))
        elif str(sh.cell_value(length, i)).strip() == 'CO4':
            if sh.cell_value(length-3, i) != '':
                c4.append(sh.cell_value(length-3, i))
        elif str(sh.cell_value(length, i)).strip() == 'CO5':
            if sh.cell_value(length-3, i) != '':
                c5.append(sh.cell_value(length-3, i))
        elif str(sh.cell_value(length, i)).strip() == 'CO6':
            if sh.cell_value(length-3, i) != '':
                c6.append(sh.cell_value(length-3, i))

    s.write(length+1, 0, '60% of CO1')
    s.write(length+2, 0, '60% of CO2')
    s.write(length+3, 0, '60% of CO3')
    s.write(length+4, 0, '60% of CO4')
    s.write(length+5, 0, '60% of CO5')
    s.write(length+6, 0, '60% of CO6')

    if len(c1) != 0:
        s.write(length+1, 1, sum(c1)/len(c1))
        if sum(c1)/len(c1) > 70:
            s.write(length+1, 2, 3)
        elif sum(c1)/len(c1) > 60:
            s.write(length+1, 2, 2)
        else:
            s.write(length+1, 2, 1)
    else:
        s.write(length+1, 1, 0)
        s.write(length+1, 2, 0)

    if len(c2) != 0:
        s.write(length+2, 1, sum(c2)/len(c2))
        if sum(c2)/len(c2) > 70:
            s.write(length+2, 2, 3)
        elif sum(c2)/len(c2) > 60:
            s.write(length+2, 2, 2)
        else:
            s.write(length+2, 2, 1)
    else:
        s.write(length+2, 1, 0)
        s.write(length+2, 2, 0)

    if len(c3) != 0:
        s.write(length+3, 1, sum(c3)/len(c3))
        if sum(c3)/len(c3) > 70:
            s.write(length+3, 2, 3)
        elif sum(c3)/len(c3) > 60:
            s.write(length+3, 2, 2)
        else:
            s.write(length+3, 2, 1)
    else:
        s.write(length+3, 1, 0)
        s.write(length+3, 2, 0)

    if len(c4) != 0:
        s.write(length+4, 1, sum(c4)/len(c4))
        if sum(c4)/len(c4) > 70:
            s.write(length+4, 2, 3)
        elif sum(c4)/len(c4) > 60:
            s.write(length+4, 2, 2)
        else:
            s.write(length+4, 2, 1)
    else:
        s.write(length+4, 1, 0)
        s.write(length+4, 2, 0)

    if len(c5) != 0:
        s.write(length+5, 1, sum(c5)/len(c5))
        if sum(c5)/len(c5) > 70:
            s.write(length+5, 2, 3)
        elif sum(c5)/len(c5) > 60:
            s.write(length+5, 2, 2)
        else:
            s.write(length+5, 2, 1)
    else:
        s.write(length+5, 1, 0)
        s.write(length+5, 2, 0)

    if len(c6) != 0:
        s.write(length+6, 1, sum(c6)/len(c6))
        if sum(c6)/len(c6) > 70:
            s.write(length+6, 2, 3)
        elif sum(c6)/len(c6) > 60:
            s.write(length+6, 2, 2)
        else:
            s.write(length+6, 2, 1)
    else:
        s.write(length+6, 1, 0)
        s.write(length+6, 2, 0)
    wb.save(excel_file_name)
    print("\nWriting Co for "+str(excel_file_name)+" ...\n")


file_names = []
for i in os.listdir():
    if i.startswith("With"):
        file_names.append(i)
subject_names = []
for i in file_names:
    calc_CO(i)
    name = ""
    count = 0
    for j in i[::-1]:
        if j == " ":
            count += 1
        if count == 2:
            if name[::-1][:-4] not in subject_names:
                subject_names.append(name[::-1][:-4])
            break
        name += j


grouped = []
for i in subject_names:
    temp = []
    for j in file_names:
        if i in j:
            if j not in temp:
                temp.append(j)
    grouped.append(temp)


def create_CO_Sheet(excel_file_names, subject_names):
    c1, c2, c3, c4, c5, c6 = [], [], [], [], [], []
    for excel_file_name in excel_file_names:
        import xlrd
        from xlrd import open_workbook
        from xlutils.copy import copy
        from random import randint
        rb = open_workbook(excel_file_name)
        sh = rb.sheet_by_index(0)
        wb = copy(rb)
        s = wb.get_sheet(0)

        try:
            if int(sh.cell_value(sh.nrows-1, 1)) > 0:
                c6.append(sh.cell_value(sh.nrows-1, 1))
        except:
            print("Error occured cannot convert to int",
                  sh.cell_value(sh.nrows-1, 1), excel_file_name)
        try:
            if int(sh.cell_value(sh.nrows-2, 1)) > 0:
                c5.append(sh.cell_value(sh.nrows-2, 1))
        except:
            print("Error occured cannot convert to int",
                  sh.cell_value(sh.nrows-2, 1), excel_file_name)
        try:

            if int(sh.cell_value(sh.nrows-3, 1)) > 0:
                c4.append(sh.cell_value(sh.nrows-3, 1))
        except:
            print("Error occured cannot convert to int",
                  sh.cell_value(sh.nrows-3, 1), excel_file_name)
        try:

            if int(sh.cell_value(sh.nrows-4, 1)) > 0:
                c3.append(sh.cell_value(sh.nrows-4, 1))
        except:
            print("Error occured cannot convert to int",
                  sh.cell_value(sh.nrows-4, 1), excel_file_name)
        try:

            if int(sh.cell_value(sh.nrows-5, 1)) > 0:
                c2.append(sh.cell_value(sh.nrows-5, 1))
        except:
            print("Error occured cannot convert to int",
                  sh.cell_value(sh.nrows-5, 1), excel_file_name)
        try:

            if int(sh.cell_value(sh.nrows-6, 1)) > 0:
                c1.append(sh.cell_value(sh.nrows-6, 1))
        except:
            print("Error occured cannot convert to int",
                  sh.cell_value(sh.nrows-6, 1), excel_file_name)

    return_arr = []
    if len(c1) > 0:
        return_arr.append(sum(c1)/len(c1))
    else:
        return_arr.append("No")
    if len(c2) > 0:
        return_arr.append(sum(c2)/len(c2))
    else:
        return_arr.append("No")
    if len(c3) > 0:
        return_arr.append(sum(c3)/len(c3))
    else:
        return_arr.append("No")
    if len(c4) > 0:
        return_arr.append(sum(c4)/len(c4))
    else:
        return_arr.append("No")
    if len(c5) > 0:
        return_arr.append(sum(c5)/len(c5))
    else:
        return_arr.append("No")
    if len(c6) > 0:
        return_arr.append(sum(c6)/len(c6))
    else:
        return_arr.append("No")
    return return_arr



def get_ass(subject_name):
    subject_name = subject_name.replace(" ","/")
    df_CO = pd.read_excel(assna,sheet_name="Sheet1")
    df_CO.head()
    index_of_subject = 0
    for i in range(len(df_CO)):
        if str(df_CO.iloc[i,0]) == subject_name:
            index_of_subject = i
    assc1,assc3 = 0,0
    assc1 = df_CO.iloc[index_of_subject+1,2]
    assc3 = df_CO.iloc[index_of_subject+2,2]
    return (assc1,assc3)



def write_final_co(subject_name):
    subject_name = subject_name.replace(" ","/")
    df_CO = pd.read_excel("Course End Survey.xlsx",sheet_name="Sheet1")
    df_CO.head()
    index_of_subject = 0
    for i in range(len(df_CO)):
        if str(df_CO.iloc[i,0]) == subject_name:
            index_of_subject = i
    arr = []
    for i in range(1,7):
        if not pd.isna(df_CO.iloc[index_of_subject+i,1]):
            arr.append( df_CO.iloc[index_of_subject+i,1] )
        else:
            break
    return arr



import xlwt 
from xlwt import Workbook 
import os
univList = []
for i in os.listdir():
    if "UNIV" in i:
        univList.append(i)


for i in grouped:
    co_arr = create_CO_Sheet(i,subject_names)
    sub_name = subject_names[grouped.index(i)]
    assc1,assc3 = get_ass(sub_name)
    
    print("                                     CO1",
          "CO2","CO3", "CO4", "CO5", "CO6",
          sep = " "*(len(str(co_arr[0]))-3)
         )
    print("Before Assignment marks -->",*co_arr,sep=" ")
    co_arr[0] = (co_arr[0]+assc1)/2
    co_arr[2] = (co_arr[2]+assc3)/2
    print("After Assignment marks -->",*co_arr,sep=" ")
    wb = Workbook() 

    sheet1 = wb.add_sheet('Sheet 1') 
    
    titles = ['CO#','Internal','Attainment Level','University',
              'Attainment','Direct CO Attainment','Direct CO Attainment Level'
              ,'Indirect CO Attainment','Indirect CO Attainment Level',
              'Final CO Attainment','Final CO Attainment Level']
    
    final_co_arr = write_final_co(sub_name)
    ctr = 0
    for i in range(len(final_co_arr)):
        sheet1.write(i+1,7,final_co_arr[ctr])
        if final_co_arr[ctr] >= 60:
            sheet1.write(i+1,8,3)
        elif final_co_arr[ctr] >= 50:
            sheet1.write(i+1,8,2)
        elif final_co_arr[ctr] >= 40:
            sheet1.write(i+1,8,1)
        else:
            sheet1.write(i+1,8,0)
        ctr+=1

    for col in range(0,len(titles)):
        sheet1.write(0,col,titles[col])
    
    str1 = "CO" 
    for j in range(1,7):
        if co_arr[j-1]=="No":
            break
        str2 = str1+str(j)
        sheet1.write(j,0,str2)
        sheet1.write(j,1,co_arr[j-1])
        if int(co_arr[j-1]) >= 60:
            sheet1.write(j,2,3)
        elif int(co_arr[j-1]) >= 50:
            sheet1.write(j,2,2)
        elif int(co_arr[j-1]) >= 40:
            sheet1.write(j,2,1)
        else:
            sheet1.write(j,2,0)
    
    
    
    
    wb.save(sub_name+".xls")
    print("\nWriting Co Calculations for "+str(sub_name)+" ...\n")




#!pip install pandas
#!pip install xlrd==1.2.0
#!pip install xlutils
#!pip install import_ipynb
import import_ipynb
# import CO_Calc
import pandas as pd
import os




def write_co_attainment(sub_name,attainment):
    excel_name = sub_name.replace('/',' ')+'.xls'
    df_Paper = pd.read_excel(excel_name)
    total_length = len(df_Paper.loc[:,])
    internal = []
    internal_att = []
    univ = []
    univ_att =[]
    for i in df_Paper.iloc[:,1]:
        internal.append(i)
    for i in df_Paper.iloc[:,2]:
        internal_att.append(i)
        univ.append(attainment)
        if attainment >= 60:
            univ_att.append(3)
        elif attainment >= 50:
            univ_att.append(2)
        elif attainment >= 40:
            univ_att.append(1)
        else:
            univ_att.append(0)

    from xlrd import open_workbook
    from xlutils.copy import copy
    rb = open_workbook(excel_name)
    wb = copy(rb)
    s = wb.get_sheet(0)
    for i in range(total_length):
        if attainment >= 60:
            s.write(i+1,4,3)
        elif attainment >= 50:
            s.write(i+1,4,2)
        elif attainment >= 40:
            s.write(i+1,4,1)
        else:
            s.write(i+1,4,0)
        s.write(i+1,3,attainment)
    for i in range(total_length):
        s.write(i+1,5,float(str((internal[i]*0.5)+(univ[i]*0.5))[:5]))
    for i in range(total_length):
        s.write(i+1,6,float(str((internal_att[i]*0.5)+(univ_att[i]*0.5))[:5]))
    wb.save(excel_name)
    





def get_sub_marks(excel_file_name):
    regulation = int(input("\nEnter regulation 2017 or 2013 for "+str(excel_file_name)+"...\n"))
    df_Paper = pd.read_excel(excel_file_name)

    column_names = df_Paper.columns[:]    
    counter = 0
    for col_name in column_names:
        grades = []
        calc = 0
        breakable=0
        for i in df_Paper.loc[1:,col_name]:
            breakable+=1
            if pd.isna(i) or i == '0' or i==0 or i==0.0 or (col_name!='S.NO' and counter == breakable):
                sub_name = col_name
                if col_name in column_names[3:]:
                    more_than_60 = 0
                    if regulation == 2017:
                        for i in grades:
                            temp = str(i).strip()
                            if temp in "AA+BB+O":
                                more_than_60+=1
                    else:
                        for i in grades:
                            temp = str(i).strip()
                            if temp in "SABCD":
                                more_than_60+=1   
                    if not sub_name.startswith("Unnamed"):
                        write_co_attainment(sub_name, more_than_60 / (counter-1) * 100)
                break
            if col_name in column_names[3:]:
                grades.append(i)
            if col_name == 'S.NO':
                counter+=1




myList = []
for i in os.listdir():
    if "UNIV" in i:
        print(i)
        get_sub_marks(i)
        print("\nCalculating university marks for "+str(i)+" ...\n")



import os
import pandas as pd
dir_percent = int(input("Enter Direct attainment percentage "))
indir_percent = int(input("Enter Indirect attainment percentage "))
subject_files = []
matrix_file_name = ""
for i in os.listdir():
    if "CO PO PSO" in i:
        matrix_file_name  = i
    if "Course End" not in i and "Assignment" not in i and "UNIV" not in i and "MODEL" not in i and "FIAT" not in i and "SIAT" not in i and ".xls" in i and len(i)<25:
        subject_files.append(i)




def write_final_co(excel_file_name):
    df_Paper = pd.read_excel(excel_file_name)
    from xlrd import open_workbook
    from xlutils.copy import copy
    rb = open_workbook(excel_file_name)
    wb = copy(rb)
    s = wb.get_sheet(0)
    length = len(df_Paper.loc[:,'Direct CO Attainment Level'])
    dir_co_att,indir_co_att = [],[]
    for i in df_Paper.loc[:,'Direct CO Attainment']:
        if pd.isna(i):
            break
        dir_co_att.append(i)
    for i in df_Paper.loc[:,'Indirect CO Attainment']:
        if pd.isna(i):
            break
        indir_co_att.append(i)
    final_co_att = [ ((dir_co_att[i])*dir_percent/100) + ((indir_co_att[i])*indir_percent/100)
                        for i in range(len(dir_co_att))]
    
    for i in range(len(dir_co_att)):
        s.write(i+1,9,final_co_att[i])
    ctr = 0
    for i in range(len(dir_co_att)):
        if final_co_att[ctr] >= 60:
            s.write(i+1,10,3)
        elif final_co_att[ctr] >= 50:
            s.write(i+1,10,2)
        elif final_co_att[ctr] >= 40:
            s.write(i+1,10,1)
        else:
            s.write(i+1,10,0)
        ctr+=1
    
    wb.save(excel_file_name)
    



def get_CO_mapping(subject_name):
    df_CO = pd.read_excel(matrix_file_name,sheet_name="Sheet1")
    df_CO.head()
    index_of_subject = 0
    for i in range(len(df_CO)):
        if str(df_CO.iloc[i,0]) == subject_name:
            index_of_subject = i
            break
    whole_List = []
    for ctr in range(1,10):
        CO_List = []
        try:
            for i in df_CO.iloc[index_of_subject+ctr,:]:
                if pd.isna(i):
                    break
                CO_List.append(i)
        except:
            pass
        if len(CO_List) <= 1 :break
        whole_List.append(CO_List)
    return whole_List




def get_list(k,check_list,po_matrix):
    temp = []
    for i in range(len(po_matrix)):
        for j in range(len(po_matrix[0])):
            if j==k:
                temp.append(po_matrix[i][j+1])
    return temp




def write_co_tabel(excel_file_name,po_matrix,check_list):
    df_Paper = pd.read_excel(excel_file_name)
    total_length = len(df_Paper.loc[:,])
    from xlrd import open_workbook
    from xlutils.copy import copy
    rb = open_workbook(excel_file_name)
    wb = copy(rb)
    s = wb.get_sheet(0)
    length = len(df_Paper.loc[:,'Direct CO Attainment Level'])
    co_att = []
    length = 0
    for i in df_Paper.loc[:,'Final CO Attainment Level']:
        if pd.isna(i):
            break
        co_att.append(i)
        length+=1
    s.write(total_length+2,0,'CO#')
    s.write(total_length+2,1,'CO Attainment')
    s.write(total_length+5+length,0,'CO#')
    s.write(total_length+5+length,1,'CO Attainment')
    for i in range(length):
        s.write(total_length+3+i,0,'CO'+str(i+1))
        s.write(total_length+3+i,1,co_att[i])
        s.write(total_length+6+i+length,0,'CO'+str(i+1))
        s.write(total_length+6+i+length,1,co_att[i])
    counter = 2
    for i in range(len(check_list[:-3])):
        if check_list[i] == 1:
            s.write(total_length+2,counter,"Weightage")
            s.write(total_length+2,counter+1,"PO"+str(i+1)+" Attainment")
            counter+=2
    counter = 2
    for i in range(len(check_list[-3:])):
        if check_list[i+len(check_list[:-3])] == 1:
            s.write(total_length+5+length,counter,"Weightage")
            s.write(total_length+5+length,counter+1,"PSO"+str(i+1)+" Attainment")
            counter+=2
    col = 2
    row_adjustment = 0
    if total_length == 5:
        row_adjustment = 1
    for k in range(len(check_list[:-3])):
        if check_list[k] == 1:        
            recieved_list = get_list(k,check_list,po_matrix)
            ctr = 0
            temp_list = []
            last_num = 0
            for i in range(total_length+2,total_length+2+len(co_att)):
                if recieved_list[ctr] == '-':
                    s.write(total_length-5+i+row_adjustment,col,'-')
                    s.write(total_length-5+i+row_adjustment,col+1,"-")
                elif recieved_list[ctr]>=3:
                    s.write(total_length-5+i+row_adjustment,col,1)
                    s.write(total_length-5+i+row_adjustment,col+1,co_att[ctr]*1)
                    temp_list.append(co_att[ctr])
                elif recieved_list[ctr]>=2:
                    s.write(total_length-5+i+row_adjustment,col,0.75)
                    s.write(total_length-5+i+row_adjustment,col+1,co_att[ctr]*0.75)
                    temp_list.append(co_att[ctr]*0.75)
                elif recieved_list[ctr] >= 1:
                    s.write(total_length-5+i+row_adjustment,col,0.5)
                    s.write(total_length-5+i+row_adjustment,col+1,co_att[ctr]*0.5)
                    temp_list.append(co_att[ctr]*0.5)
                else:
                    s.write(total_length-5+i+row_adjustment,col,0)
                    s.write(total_length-5+i+row_adjustment,col+1,0)
                    temp_list.append(0)
                ctr+=1
                last_num = i
            s.write(total_length-5+last_num+1+row_adjustment,col+1,sum(temp_list)/len(temp_list))
            col+=2
        
    col = 2
    row_adjustment = 0
    if total_length == 5:
        row_adjustment = 1
    for k in range(len(check_list[-3:])):
        if check_list[k+len(check_list[:-3])] == 1:        
            recieved_list = get_list(k+len(check_list[:-3]),check_list,po_matrix)
            ctr = 0
            temp_list = []
            last_num = 0
            for i in range(total_length+5+length,total_length+5+len(co_att)+length):
                if recieved_list[ctr] == '-':
                    s.write(total_length-5+i+row_adjustment,col,'-')
                    s.write(total_length-5+i+row_adjustment,col+1,"-")
                elif recieved_list[ctr]>=3:
                    s.write(total_length-5+i+row_adjustment,col,1)
                    s.write(total_length-5+i+row_adjustment,col+1,co_att[ctr]*1)
                    temp_list.append(co_att[ctr])
                elif recieved_list[ctr]>=2:
                    s.write(total_length-5+i+row_adjustment,col,0.75)
                    s.write(total_length-5+i+row_adjustment,col+1,co_att[ctr]*0.75)
                    temp_list.append(co_att[ctr]*0.75)
                elif recieved_list[ctr] >= 1:
                    s.write(total_length-5+i+row_adjustment,col,0.5)
                    s.write(total_length-5+i+row_adjustment,col+1,co_att[ctr]*0.5)
                    temp_list.append(co_att[ctr]*0.5)
                else:
                    s.write(total_length-5+i+row_adjustment,col,0)
                    s.write(total_length-5+i+row_adjustment,col+1,0)
                    temp_list.append(0)
                ctr+=1
                last_num = i
            s.write(total_length-5+last_num+1+row_adjustment,col+1,sum(temp_list)/len(temp_list))
            col+=2
            
            
    wb.save(excel_file_name)





for i in subject_files:
    write_final_co(i)
    try:
        excel_file_name = i
        temp = i.replace(" ","/")[:-4]
        po_matrix = get_CO_mapping(temp)
        print(temp,end="\n\n")
        print("      PO1  PO2  PO3  PO4  PO5  PO6  PO7  PO8  PO9 PO10 PO11 PO12 PSO1 PSO2  PSO")
        for row in po_matrix:
            print(*row,sep="    ")
        check_list = [0 for i in range(len(po_matrix[0]))]
        for i in range(len(po_matrix)):
            for j in range(1,len(po_matrix[i])):
                if po_matrix[i][j]!='-':
                    check_list[j] = 1
        print("   ",*check_list[1:],sep="    ")
        print("\n--------------------------------------------------------------------------------\n")
        write_co_tabel(excel_file_name,po_matrix,check_list[1:])
    except:
        print("\n\nSkipping "+str(temp)+"...\n\nSome Error occured...\n\nPlease check the file Manually...\n\n")
        print("Data Skipped")
        print(excel_file_name)
        print(po_matrix,sep="\n")
        print(check_list[1:])





import shutil
output_files = []
try:
    print("\n\nTrying to create Output folder...\n\n")
    os.makedirs("Output")
    print("\n\nOutput Folder created successfully...\n\n")
except:
    print("\n\nOutput folder already exist...\n\n")
for i in os.listdir():
    if ".xls" in i and ".xlsx" not in i:
        output_files.append(i)
for i in output_files:
    print("\nMoving "+str(i)+" to Outputs folder")
    shutil.move(os.path.basename(i),"Output/"+str(os.path.basename(i)))



print("\nProgram Executed Successfully...\n\nDownload Your Files from Output folder...\n")


