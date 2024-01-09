from openpyxl import Workbook
from openpyxl import load_workbook
import os

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
        if ws['B'+str(cell.row)].value == None:
            continue
            
        if ws['A'+str(cell.row)].value == None:
            f1.write("ERROR: Row-"+str(cell.row) + ": Name is Empty\n")
        if len(str(ws['B'+str(cell.row)].value)) < 5:
            f2.write("ERROR: Row-"+str(cell.row) + ": Sign On Name String is less than 5\n")
        if len(str(ws['D'+str(cell.row)].value)) > 2:
            f3.write("WARNING Row-"+str(cell.row) + ": DAO Warning (Greater than 2)\n")
        if ws['E'+str(cell.row)].value == None or len(str(ws['E'+str(cell.row)].value)) != 9:
            f4.write("ERROR: Row-"+str(cell.row) + ": Primary Company is Empty or not length of 9\n")
        if ws['C'+str(cell.row)].value == None:
            f5.write("ERROR: Row-"+str(cell.row) + ": Role is Field\n")

        nws['A'+str(cell.row)] = ""
        nws['B'+str(cell.row)] = "U."+str(ws['B'+str(cell.row)].value).rjust(5,"0") if ws['B'+str(cell.row)].value else ""
        nws['C'+str(cell.row)] = str(ws['A'+str(cell.row)].value).upper().strip() if ws['A'+str(cell.row)].value else ""
        nws['D'+str(cell.row)] = str(ws['B'+str(cell.row)].value).rjust(5,"0") if ws['B'+str(cell.row)].value else ""
        nws['E'+str(cell.row)] = "INT"
        nws['F'+str(cell.row)] = "1"
        nws['H'+str(cell.row)] = ws['D'+str(cell.row)].value
        nws['I'+str(cell.row)] = "20240101M0101"
        nws['J'+str(cell.row)] = "20231218"
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
               
        appl = "ALL.PG" #Other than checker
        c_appl = "@ASA.PK.A.CAD" #Checker
        m_func = "B C D E F H I L P R S V" #Maker
        c_func = "" #Checker
        v_func = "H L P S V" #View Only
        a_func = "A 2 B C D E F H I L P R S V" #All Rights
        
        P_COMPANY = ws['E'+str(cell.row)].value.replace(" ","")
        P_COMPANY = P_COMPANY.replace(",","")
        P_COMPANY = "ALL::"+P_COMPANY
        
        if ws['F'+str(cell.row)].value != None:
            COMPANY = ws['F'+str(cell.row)].value.replace(" ","")
            COMPANY = COMPANY.replace(",,",",")
            COMPANY = COMPANY.replace(",","::")
            COMPANY = P_COMPANY+"::"+COMPANY
        else:
            COMPANY = P_COMPANY
        
        items = COMPANY.split("::")
        items = [x for n, x in enumerate(items) if x not in items[:n]]
        items = [x for n, x in enumerate(items) if len(x) == 9 or x == "ALL"]
        
        COMPANY = "::".join(items)
        
        xlen = len(items)
        application = ""
        function = ""
        for x in range(xlen):
            if x < 2:
                if ws['C'+str(cell.row)].value == "Maker":
                    application += appl+"::"
                    function    += m_func+"::"
                elif ws['C'+str(cell.row)].value == "Checker":
                    application += c_appl+"::"
                    function    += c_func+"::"
                elif ws['C'+str(cell.row)].value == "View":
                    application += appl+"::"
                    function    += v_func+"::"
                elif ws['C'+str(cell.row)].value == "ALL":
                    application += appl+"::"
                    function    += a_func+"::"
            else:
                application +=  appl+"::"
                function += v_func+"::"
                
        application = application[0:len(application)-2]
        function = function[0:len(function)-2]
            
        if xlen > 2:
            COMPANY = COMPANY.replace("ALL::","")
            
        nws['G'+str(cell.row)] = COMPANY
        
        if xlen > 2:
            nws['P'+str(cell.row)] = COMPANY + "::ALL"
            
            t_application = application.split("::")
            t_function = function.split("::")
            
            t_application = t_application[1:] + t_application[:1]
            t_function = t_function[1:] + t_function[:1]
            
            application = "::".join(t_application)
            function = "::".join(function)
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
        