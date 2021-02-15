import random
f = open("Input.txt", "r")
marks = f.readlines()
markList = [i.strip() for i in marks]
print(markList)
def splitMarks(ans,marksPartC,marksPartB,marksPartA):
    originalMarks = ans
    ans = int(ans)
    
    ##Part C Marks Allocation
    for i in range(1):
        partCmark = random.randint(0,15)
        if partCmark+marksPartC[0]>15:
            break
        if partCmark>=ans:
            ans = partCmark
            return (marksPartC,marksPartB,marksPartA)
        marksPartC[0] += partCmark
        ans-=partCmark
    
    
    ##Part B Marks Allocation
    partBmark = random.randint(0,13)
    for i in range(5):
        if partBmark+sum(marksPartB)>65:
            break
        if(ans-partBmark < 0):
            partBmark = ans
            break
        marksPartB[i] += partBmark
        ans-=partBmark
    
    ##Part A Marks Allocation
    for i in range(10):
        partAmark = random.randint(0,2)
        if partAmark+sum(marksPartA)>20:
            break
        if(ans-partAmark < 0):
            partBmark = ans
            break
        marksPartA[i] += partAmark
        ans-=partAmark
    return (marksPartA,marksPartB,marksPartC)
def adjustMarks(temp,a,b,c):
    while True:
        for i in range(1):
            if temp==0:
                return (a,b,c)
            if c[i]<16:
                c[i]+=1
                temp-=1
        for j in range(5):
            if temp==0:
                return (a,b,c)
            if b[j]<13:
                b[j]+=1
                temp-=1
        for k in range(10):
            if temp==0:
                return (a,b,c)
            if a[k] < 2:
                a[k]+=1
                temp-=1
    return (a,b,c)
writeToExcel = []

for i in markList:
    marksPartC = [0]
    marksPartB = [0 for i in range(5)]
    marksPartA = [0 for i in range(10)]
    
    if i.isalpha():continue ##Handling Absenties
        
    (a,b,c) = splitMarks(i,marksPartC,marksPartB,marksPartA)
    if sum(a+b+c)<int(i):
        temp = int(i) - sum(a+b+c)
        (a,b,c) = adjustMarks(temp,a,b,c)
    writeToExcel.append(a+b+c+[int(i)])

temp=[]
for i in range(len(writeToExcel)):
    single = []
    for j in range(10):
        single.append(writeToExcel[i][j])
    for j in range(10,15):
        aORb = random.randint(0,1)
        if aORb == 0:
            single.append(0)
            single.append(writeToExcel[i][j])
        else:
            single.append(writeToExcel[i][j])
            single.append(0)
    for j in range(15,len(writeToExcel[i])):
        single.append(writeToExcel[i][j])
    temp.append(single)
    
writeToExcel = temp

import xlwt 
from xlwt import Workbook 

wb = Workbook() 
  
sheet1 = wb.add_sheet('Sheet 1') 

for i in range(10):
    sheet1.write(0,i+1,i+1)
counter = 11
for i in range(11,16):
    sheet1.write(0,counter,i)
    sheet1.write(0,counter+1,i)
    counter+=2
sheet1.write(0,counter,16)
for i in range(len(temp)):
    sheet1.write(i+1,0,"names")

for i in range(len(temp)):
    for j in range(len(temp[i])):
        sheet1.write(i+1,j+1,int(temp[i][j]))
  
wb.save('marks.xls')