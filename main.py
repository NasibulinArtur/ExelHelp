from thefuzz import fuzz
import openpyxl as openpyxl


def sim(s1,s2):
    normalized = s1.lower()
    ch = ")"
    if(ch in normalized):
        normalized1 = normalized.split(')', 1)[1]
    else: normalized1 = normalized
    normalized2 = s2.lower()
    matcher = fuzz.token_sort_ratio(normalized1,normalized2)
    return matcher

exel = []



wb = openpyxl.load_workbook(filename='MVA.xlsx')
sheet = wb["Прайс-Лист"]



max = 60
max_value = []


citys = ["Уфа"]
matrix = []
matrix2 =[]
for city in citys:
    #качаю ексель таблицу с данными
    wb1 = openpyxl.load_workbook(filename='Красноярск.xlsx')
    sheet1 = wb1.active
    for row in sheet1.iter_rows(min_row=13,values_only=True):
        rowlist = list(row)
        if (rowlist[13] == None):
            rowlist[13] = 0.0
        if (rowlist[14] == None):
            rowlist[14] = 0.0
        if (rowlist[15] == None):
            rowlist[15] = 0.0
        if (rowlist[13] >=1000):
            rowlist[13] = rowlist[13]/1000
        if (rowlist[14] >=1000):
            rowlist[14] = rowlist[14]/1000
        if (rowlist[15] >=1000):
            rowlist[15] = rowlist[15]/1000
        if(rowlist[14] + rowlist[15]<rowlist[13]+3):
            need_to_buy = rowlist[13] - rowlist[14] - rowlist[15] + 2
            matrix.append([rowlist[0], need_to_buy])


for nom in matrix:
    for row in sheet.iter_rows(values_only=True):
        if (sim(nom[0],str(row[0])) > max):
            max = sim(nom[0],str(row[0]))
            try:
                max_value = [str(nom[0]),row[0], sim(nom[0],str(row[0])) , str(row[8]).split('*')[1]]
            except:
                max_value = [str(nom[0]),row[0], sim(nom[0],str(row[0])) , "0"]
        if (str(row[0]) == "Процессор Intel Celeron G4930 Soc-1151v2 (3.2GHz/iUHDG610) OEM") :
            if max_value != '':
                matrix2.append([max_value])
            max_value = ""
            max = 60
            break

for i in matrix2:
    print(i)
    a = input()
    if(a == "1"):
        sheet[i[0][3]].value = 5
        print(f"Записал")
    if (a == "0"):
        print(f"Не записал")

wb.save('balances.xlsx')