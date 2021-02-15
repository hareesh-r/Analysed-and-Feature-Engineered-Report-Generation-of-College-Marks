print("Enter 1 for IA1 2 for IA2 else open required file")
inp = input()
if inp == "1":
    first = "CO1"
    second = "CO2"
elif inp == "2":
    first = "CO3"
    second = "CO4"
import xlrd

loc = ("marks.xls")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

Q1,Q2,Q3,Q4,Q5,Q6,Q7,Q8,Q9,Q10,Q11,Q12,Q13,Q14,Q151,Q152,Q161,Q162 = 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P151,P152,P161,P162 = 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0

for i in range(1,sheet.nrows):
    for j in range(sheet.ncols):
        try:
            int(sheet.cell_value(i,j))
            if sheet.cell_value(i,j)>0:
                if j==1:
                    Q1+=1
                elif j==2:
                    Q2+=1
                elif j==3:
                    Q3+=1
                elif j==4:
                    Q4+=1
                elif j==5:
                    Q5+=1
                elif j==6:
                    Q6+=1
                elif j==7:
                    Q7+=1
                elif j==8:
                    Q8+=1
                elif j==9:
                    Q9+=1
                elif j==10:
                    Q10+=1
                elif (j==11 or j==12):
                    Q11+=1
                elif (j==13 or j==14):
                    Q12+=1
                elif (j==15 or j==16):
                    Q13+=1
                elif (j==17 or j==18):
                    Q14+=1
                elif j==19:
                    Q151+=1
                elif j==20:
                    Q152+=1
                elif j==21:
                    Q161+=1
                elif j==22:
                    Q162+=1
            if sheet.cell_value(i,j)>1.2:
                if j==1:
                    P1+=1
                elif j==2:
                    P2+=1
                elif j==3:
                    P3+=1
                elif j==4:
                    P4+=1
                elif j==5:
                    P5+=1
                elif j==6:
                    P6+=1
                elif j==7:
                    P7+=1
                elif j==8:
                    P8+=1
                elif j==9:
                    P9+=1
                elif j==10:
                    P10+=1
            if sheet.cell_value(i,j)>7.8:
                if (j==11 or j==12):
                    P11+=1
                elif (j==13 or j==14):
                    P12+=1
                elif (j==15 or j==16):
                    P13+=1
                elif (j==17 or j==18):
                    P14+=1
                elif j==19:
                    P151+=1
                elif j==20:
                    P152+=1
            if sheet.cell_value(i,j)>9:
                if j==21:
                    P161+=1
                elif j==22:
                    P162+=1
        except:
            pass
print(Q1,Q2,Q3,Q4,Q5,Q6,Q7,Q8,Q9,Q10,Q11,Q12,Q13,Q14,Q151,Q152,Q161,Q162)
print(P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P151,P152,P161,P162)

import xlwt 
from xlwt import Workbook 

wb = Workbook() 

if inp == "1":
    sheet1 = wb.add_sheet('IA 1') 
elif inp == "2":
    sheet1 = wb.add_sheet('IA 2') 

for i in range(sheet.nrows):
    for j in range(sheet.ncols):
        sheet1.write(i,j,sheet.cell_value(i, j))
        
i+=3

counter=1
for col in range(1,6):
    sheet1.write(i,col,r"IA-"+str(inp)+"-Q"+str(counter)+r"(2 MARKS)")
    counter+=1
sheet1.write(i,col+1,r"IA-"+str(inp)+"-Q11 (13 MARKS)")
sheet1.write(i,col+2,r"IA-"+str(inp)+"-Q12 (13 MARKS)")
sheet1.write(i,col+3,r"IA-"+str(inp)+"-Q15A (13 MARKS)")
sheet1.write(i,col+4,r"IA-"+str(inp)+"-Q16A (15 MARKS)")
counter+=4
for col2 in range(col+5,col+10):
    sheet1.write(i,col2,r"IA-"+str(inp)+"-Q"+str(counter)+r"(2 MARKS)")
    counter+=1
sheet1.write(i,col2+1,r"IA-"+str(inp)+"-Q13 (13 MARKS)")
sheet1.write(i,col2+2,r"IA-"+str(inp)+"-Q14 (13 MARKS)")
sheet1.write(i,col2+3,r"IA-"+str(inp)+"-Q15B (13 MARKS)")
sheet1.write(i,col2+4,r"IA-"+str(inp)+"-Q16B (15 MARKS)")
counter+=3
sheet1.write(i+3,1,first)
sheet1.write(i+3,11,second)
sheet1.write(i+4,0,"Number of students attempted")
sheet1.write(i+5,0,"Number of Students Scored >= 60% of Marks")
sheet1.write(i+6,0,"Percentage of Students Scored >= 60% of Marks")
sheet1.write(i+7,0,"Average number of students scored >=60% marks")
sheet1.write(i+8,0,"CO Attainment Level")

for j in range(1,sheet.ncols-4):
    if j==1:
        sheet1.write(i+4,j,Q1)
    elif j==2:
        sheet1.write(i+4,j,Q2)
    elif j==3:
        sheet1.write(i+4,j,Q3)
    elif j==4:
        sheet1.write(i+4,j,Q4)
    elif j==5:
        sheet1.write(i+4,j,Q5)
    elif j==6:
        sheet1.write(i+4,j,Q11)
    elif j==7:
        sheet1.write(i+4,j,Q12)
    elif j==8:
        sheet1.write(i+4,j,Q151)
    elif j==9:
        sheet1.write(i+4,j,Q161)
    elif j==10:
        sheet1.write(i+4,j,Q6)
    elif j==11:
        sheet1.write(i+4,j,Q7)
    elif j==12:
        sheet1.write(i+4,j,Q8)
    elif j==13:
        sheet1.write(i+4,j,Q9)
    elif j==14:
        sheet1.write(i+4,j,Q10)
    elif j==15:
        sheet1.write(i+4,j,Q13)
    elif j==16:
        sheet1.write(i+4,j,Q14)
    elif j==17:
        sheet1.write(i+4,j,Q152)
    elif j==18:
        sheet1.write(i+4,j,Q162)
i+=1

for j in range(1,sheet.ncols-4):
    if j==1:
        sheet1.write(i+4,j,P1)
    elif j==2:
        sheet1.write(i+4,j,P2)
    elif j==3:
        sheet1.write(i+4,j,P3)
    elif j==4:
        sheet1.write(i+4,j,P4)
    elif j==5:
        sheet1.write(i+4,j,P5)
    elif j==6:
        sheet1.write(i+4,j,P11)
    elif j==7:
        sheet1.write(i+4,j,P12)
    elif j==8:
        sheet1.write(i+4,j,P151)
    elif j==9:
        sheet1.write(i+4,j,P161)
    elif j==10:
        sheet1.write(i+4,j,P6)
    elif j==11:
        sheet1.write(i+4,j,P7)
    elif j==12:
        sheet1.write(i+4,j,P8)
    elif j==13:
        sheet1.write(i+4,j,P9)
    elif j==14:
        sheet1.write(i+4,j,P10)
    elif j==15:
        sheet1.write(i+4,j,P13)
    elif j==16:
        sheet1.write(i+4,j,P14)
    elif j==17:
        sheet1.write(i+4,j,P152)
    elif j==18:
        sheet1.write(i+4,j,P162)
i+=1

for j in range(1,sheet.ncols-4):
    if j==1:
        sheet1.write(i+4,j,int(P1/Q1*100))
    elif j==2:
        sheet1.write(i+4,j,int(P2/Q2*100))
    elif j==3:
        sheet1.write(i+4,j,int(P3/Q3*100))
    elif j==4:
        sheet1.write(i+4,j,int(P4/Q4*100))
    elif j==5:
        sheet1.write(i+4,j,int(P5/Q5*100))
    elif j==6:
        sheet1.write(i+4,j,int(P11/Q11*100))
    elif j==7:
        sheet1.write(i+4,j,int(P12/Q12*100))
    elif j==8:
        sheet1.write(i+4,j,int(P151/Q151*100))
    elif j==9:
        sheet1.write(i+4,j,int(P161/Q161*100))
    elif j==10:
        sheet1.write(i+4,j,int(P6/Q6*100))
    elif j==11:
        sheet1.write(i+4,j,int(P7/Q7*100))
    elif j==12:
        sheet1.write(i+4,j,int(P8/Q8*100))
    elif j==13:
        sheet1.write(i+4,j,int(P9/Q9*100))
    elif j==14:
        sheet1.write(i+4,j,int(P10/Q10*100))
    elif j==15:
        sheet1.write(i+4,j,int(P13/Q13*100))
    elif j==16:
        sheet1.write(i+4,j,int(P14/Q14*100))
    elif j==17:
        sheet1.write(i+4,j,int(P152/Q152*100))
    elif j==18:
        sheet1.write(i+4,j,int(P162/Q162*100))
i+=1

avg1 = (int(P1/Q1*100) + int(P2/Q2*100) + int(P3/Q3*100) + int(P4/Q4*100) + int(P5/Q5*100) + int(P11/Q11*100) + int(P12/Q12*100) + int(P151/Q151*100) + int(P161/Q161*100))/9

avg2 = (int(P6/Q6*100) + int(P7/Q7*100) + int(P8/Q8*100) + int(P9/Q9*100) + int(P10/Q10*100) + int(P13/Q13*100) + int(P14/Q14*100) + int(P152/Q152*100) + int(P162/Q162*100))/9

sheet1.write(i+4,1,avg1)

sheet1.write(i+4,11,avg2)

i+=1

att = 1
if avg1>70:
    att = 3
elif avg1>60:
    att = 2
sheet1.write(i+4,1,att)

att = 1
if avg2>70:
    att = 3
elif avg2>60:
    att = 2
sheet1.write(i+4,11,att)

fileName = "Internal Assessment "+inp+".xls"
wb.save(fileName)