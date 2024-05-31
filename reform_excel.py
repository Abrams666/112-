#imports
import openpyxl

#read old excel
file_path = "112_result_school_data.xlsx"
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

#write new excel
workbook = openpyxl.Workbook()
sheetx = workbook.worksheets[0]

sheetx["A1"]="校名"
sheetx["B1"]="系名"
sheetx["C1"]="國文加權"
sheetx["D1"]="英文加權"
sheetx["E1"]="數甲加權"
sheetx["F1"]="數A加權"
sheetx["G1"]="數B加權"
sheetx["H1"]="物理加權"
sheetx["I1"]="化學加權"
sheetx["J1"]="生物加權"
sheetx["K1"]="地理加權"
sheetx["L1"]="歷史加權"
sheetx["M1"]="公民加權"
sheetx["N1"]="錄取分數"

k=1
for i in range(3,1898):
    k=k+1
    if (sheet["B"+str(i)].value=="校名"):
        k=k-1
    else:
        sheetx["A"+str(k)]=sheet["B"+str(i)].value
        sheetx["B"+str(k)]=sheet["C"+str(i)].value
        if(str(sheet["F"+str(i)].value)=="-----"):
            sheetx["N"+str(k)]=0
        else:
            sheetx["N"+str(k)]=sheet["F"+str(i)].value

        print(len(sheet["D"+str(i)].value))
    
        for j in range(len(sheet["D"+str(i)].value)-1):
            if(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="國x"):
                sheetx["C"+str(k)]=float(str(sheet["D"+str(i)].value[j+2])+str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="英x"):
                sheetx["D"+str(k)]=float(str(sheet["D"+str(i)].value[j+2])+str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="數甲"):
                sheetx["E"+str(k)]=float(str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5])+str(sheet["D"+str(i)].value[j+6]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="數A"):
                sheetx["F"+str(k)]=float(str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5])+str(sheet["D"+str(i)].value[j+6]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="數B"):
                sheetx["G"+str(k)]=float(str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5])+str(sheet["D"+str(i)].value[j+6]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="物x"):
                sheetx["H"+str(k)]=float(str(sheet["D"+str(i)].value[j+2])+str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="化x"):
                sheetx["I"+str(k)]=float(str(sheet["D"+str(i)].value[j+2])+str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="生x"):
                sheetx["J"+str(k)]=float(str(sheet["D"+str(i)].value[j+2])+str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="地x"):
                sheetx["K"+str(k)]=float(str(sheet["D"+str(i)].value[j+2])+str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="歷x"):
                sheetx["L"+str(k)]=float(str(sheet["D"+str(i)].value[j+2])+str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5]))
            elif(str(sheet["D"+str(i)].value[j])+str(sheet["D"+str(i)].value[j+1])=="公x"):
                sheetx["M"+str(k)]=float(str(sheet["D"+str(i)].value[j+2])+str(sheet["D"+str(i)].value[j+3])+str(sheet["D"+str(i)].value[j+4])+str(sheet["D"+str(i)].value[j+5]))

workbook.save('Exam_Subject_Score_Weight.xlsx')