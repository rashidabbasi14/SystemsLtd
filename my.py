from openpyxl import Workbook
from openpyxl import load_workbook
import os

class bcolors:
    WARNING = '\033[93m'
    ERROR = '\033[91m'

wb = load_workbook(filename = 'UserSheet.xlsx')
nwb = Workbook()
ws = wb.active
nws = nwb.active


mincol = 1
maxcol = 1
minrow = 2
maxrow = ws.max_row


nws['A1'] = "UPLOAD.COMPANY"
nws['B1'] = "@ID"
nws['C1'] = "USER.NAME"
nws['D1'] = "SIGN.ON.NAME"
nws['E1'] = "CLASSIFICATION"
nws['F1'] = "LANGUAGE"
nws['G1'] = "COMPANY.CODE"
nws['H1'] = "DEPARTMENT.CODE"
nws['I1'] = "PASSWORD.VALIDITY"
nws['J1'] = "START.DATE.PROFILE"
nws['K1'] = "END.DATE.PROFILE"
nws['L1'] = "START.TIME"
nws['M1'] = "END.TIME"
nws['N1'] = "TIME.OUT.MINUTES"
nws['O1'] = "ATTEMPTS"
nws['P1'] = "COMPANY.RESTR"
nws['Q1'] = "APPLICATION"
nws['R1'] = "FUNCTION"
nws['S1'] = "SIGN.ON.OFF.LOG"
nws['T1'] = "SECURITY.MGMT.L"
nws['U1'] = "APPLICATION.LOG"
nws['V1'] = "FUNCTION.ID.LOG"
nws['W1'] = "INPUT.DAY.MONTH"
nws['X1'] = "CLEAR.SCREEN"
nws['Y1'] = "DEALER.DESK"
nws['Z1'] = "AMOUNT.FORMAT"
nws['AA1'] = "DATE.FORMAT"


if not os.path.exists("Errors"):  
    os.makedirs("Errors") 
    
f1 = open("Errors/nameError.txt", "w")
f2 = open("Errors/signOnNameError.txt", "w")
f3 = open("Errors/DAO Warning.txt", "w")
f4 = open("Errors/CompanyNameEmpty.txt", "w")
f5 = open("Errors/MultipleRoleError.txt", "w")

for col in ws.iter_cols(min_row = minrow, min_col = mincol, max_col = maxcol, max_row = maxrow):
    for cell in col:
        if ws['D'+str(cell.row)].value == None:
            continue
        if ws['B'+str(cell.row)].value == None:
            print("ERROR: Row-"+str(cell.row) + ": Name is Empty")
            f1.write("ERROR: Row-"+str(cell.row) + ": Name is Empty\n")
        if len(str(ws['D'+str(cell.row)].value)) < 5:
            print("ERROR: Row-"+str(cell.row) + ": Sign On Name String is less than 5")
            f2.write("ERROR: Row-"+str(cell.row) + ": Sign On Name String is less than 5\n")
        if len(str(ws['I'+str(cell.row)].value)) > 2:
            print("WARNING Row-"+str(cell.row) + ": DAO Warning (Greater than 2)")
            f3.write("WARNING Row-"+str(cell.row) + ": DAO Warning (Greater than 2)\n")
        if ws['J'+str(cell.row)].value == None:
            print("ERROR: Row-"+str(cell.row) + ": Company is Empty")
            f4.write("ERROR: Row-"+str(cell.row) + ": Company is Empty\n")

        count = 0
        if ws['E'+str(cell.row)].value == "Y":
            count = count + 1 
        if ws['F'+str(cell.row)].value == "Y":
            count = count + 1 
        if ws['G'+str(cell.row)].value == "Y":
            count = count + 1 
        if ws['H'+str(cell.row)].value == "Y":
            count = count + 1
        if count > 1: 
            print("ERROR: Row-"+str(cell.row) + ": Multiple roles are assigned")
            f5.write("ERROR: Row-"+str(cell.row) + ": Multiple roles are assigned\n")
            
        if ws['E'+str(cell.row)].value == None or ws['F'+str(cell.row)].value == None or ws['G'+str(cell.row)].value == None or ws['H'+str(cell.row)].value == None:
            print("ERROR: Row-"+str(cell.row) + ": Empty Field")
            f5.write("ERROR: Row-"+str(cell.row) + ": Empty Field\n")
            
        if ws['E'+str(cell.row)].value == None and ws['F'+str(cell.row)].value == None and ws['G'+str(cell.row)].value == None and ws['H'+str(cell.row)].value == None:
            print("ERROR: Row-"+str(cell.row) + ": ALL Role Empty Field")
            f5.write("ERROR: Row-"+str(cell.row) + ": ALL Role Empty Field\n")

        nws['A'+str(cell.row)] = ""
        nws['B'+str(cell.row)] = "U."+str(ws['D'+str(cell.row)].value).rjust(5,"0") if ws['D'+str(cell.row)].value else ""
        nws['C'+str(cell.row)] = ws['B'+str(cell.row)].value.upper().strip() if ws['B'+str(cell.row)].value else ""
        nws['D'+str(cell.row)] = str(ws['D'+str(cell.row)].value).rjust(5,"0") if ws['D'+str(cell.row)].value else ""
        nws['E'+str(cell.row)] = "INT"
        nws['F'+str(cell.row)] = "1"
        nws['H'+str(cell.row)] = ws['I'+str(cell.row)].value
        nws['I'+str(cell.row)] = "20240101M0101"
        nws['J'+str(cell.row)] = "20231412"
        nws['K'+str(cell.row)] = "20990321"
        nws['L'+str(cell.row)] = "00:00"
        nws['M'+str(cell.row)] = "24:00"
        nws['N'+str(cell.row)] = "999"
        nws['O'+str(cell.row)] = "9"
        nws['S'+str(cell.row)] = "Y"
        nws['T'+str(cell.row)] = "Y"
        nws['U'+str(cell.row)] = "Y"
        nws['V'+str(cell.row)] = "Y"
        nws['W'+str(cell.row)] = "DDMM"
        nws['X'+str(cell.row)] = "Y"
        nws['Y'+str(cell.row)] = "00"
        nws['Z'+str(cell.row)] = "?."
        nws['AA'+str(cell.row)] = "1"

        if(ws['J'+str(cell.row)].value):
            appl = ""
            func = ""
            if ws['E'+str(cell.row)].value == "Y": #Maker
                appl = "ALL.PG"
                func = "B C D E F H I L P R S V"
            elif ws['F'+str(cell.row)].value == "Y": #Checker
                appl = "@ASA.PK.A.CAD"
                func = ""
            elif ws['G'+str(cell.row)].value == "Y": #View Only
                appl = "ALL.PG"
                func = "H L P S V"
            elif ws['H'+str(cell.row)].value == "Y": #All Rights
                appl = "ALL.PG"
                func = "A 2 B C D E F H I L P R S V"
            
            COMPANY = ws['J'+str(cell.row)].value.replace(" ","")
            COMPANY = COMPANY.replace(",,",",")
            COMPANY = COMPANY.replace(",","::")
            COMPANY = "ALL::"+COMPANY
            items = COMPANY.split("::")
            items = [x for n, x in enumerate(items) if x not in items[:n]]
            items = [x for n, x in enumerate(items) if len(x) == 9 or x == "ALL"]
            
            COMPANY = "::".join(items)
            
            xlen = len(items)
            application = appl
            function = func
            for x in range(xlen-1):
                if appl:
                    application += "::"+appl
                if func:
                    function += "::"+func

            if xlen > 2:
                COMPANY = COMPANY.replace("ALL::","")
                
            nws['G'+str(cell.row)] = COMPANY
            
            if xlen > 2:
                nws['P'+str(cell.row)] = COMPANY + "::ALL"
            else:
                nws['P'+str(cell.row)] = COMPANY
                
            nws['Q'+str(cell.row)] = application
            nws['R'+str(cell.row)] = function
            
nwb.save('DMT.xlsx')

f1.close()
f2.close()
f3.close()
f4.close()
f5.close()

files = ["Errors/nameError.txt","Errors/signOnNameError.txt","Errors/DAO Warning.txt","Errors/CompanyNameEmpty.txt","Errors/MultipleRoleError.txt"]

for file in files:
    file_stats = os.stat(file)
    if(file_stats.st_size == 0):
        os.remove(file)
        