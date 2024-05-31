#imports
import openpyxl

#defs
def cross_multiplication(a,b):
    if (len(a)>=len(b)):
        len_ab=len(b)
    else:
        len_ab=len(a)
    c=[]
    for i in range(len_ab):
        d=a[i]*b[i]
        c.append(d)
    return c

def array_plus(a):
    b=0
    for i in range(len(a)):
        b=b+a[i]
    return b

#get information
scores=[47,46,37,36,54,28,35,29,0,0,0]
print("\/請輸入各科原始成績(未報考則輸入0)")
scores[0]=float(input("國文>"))
scores[1]=float(input("英文>"))
scores[2]=float(input("數甲>"))
scores[3]=float(input("數A >"))
scores[4]=float(input("數B >"))
scores[5]=float(input("物理>"))
scores[6]=float(input("化學>"))
scores[7]=float(input("生物>"))
scores[8]=float(input("地理>"))
scores[9]=float(input("歷史>"))
scores[10]=float(input("公民>"))

#read weight
file_path = "Exam_Subject_Score_Weight.xlsx"
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

weights=[[0 for _ in range(11)] for _ in range(1838)]

for i in range(2,1840):
    for j in range(2,13):
        if(str(sheet[i][j].value)=="None"):
            weights[i-2][j-2]=0
        else:
            weights[i-2][j-2]=float(sheet[i][j].value)
    if((float(array_plus(cross_multiplication(scores,weights[i-2])))>=float(sheet["N"+str(i)].value))):# and ((str(sheet["A"+str(i)].value)[0]+str(sheet["A"+str(i)].value)[1])=="國立")
        print(str(i)+" "+str(sheet["A"+str(i)].value)+" "+str(sheet["B"+str(i)].value))