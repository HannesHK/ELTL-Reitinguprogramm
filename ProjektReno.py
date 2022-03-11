import xlrd #import excel reader
import xlwt #import excel writer

loc1 = ("ReitingAktiivne.xls") #aktiivse reitingu fail

wb1 = xlrd.open_workbook(loc1) #avab exceli workbooki
sheet1 = wb1.sheet_by_index(0) #avab workbookis esimese lehe

#sonastik koigist praegu reitingus olevatest inimestest
ID_RP = {}
for i in range(1, sheet1.nrows):              #Paigutuspunktid                 Reitingupunktid           Kaalud                        Perenimi                Eesnimi              Sugu                        Sünnipäev            ReitinguKuupäev                Klubi nimi                    
    ID_RP[int(sheet1.cell_value(i,0))] = [int(sheet1.cell_value(i,7)), int(sheet1.cell_value(i,8)), int(sheet1.cell_value(i,9)), sheet1.cell_value(i,1), sheet1.cell_value(i,2), sheet1.cell_value(i,3), str(sheet1.cell_value(i,4)), str(sheet1.cell_value(i,10)), sheet1.cell_value(i,11)]
    

    
loc2 = ("protokoll.xls") #võistluse protokolli fail

wb2 = xlrd.open_workbook(loc2) 
sheet2 = wb2.sheet_by_index(0)

#sonastik turniiril osalenutest
dict1 = {}
for j in range(8, sheet2.nrows):
    mänginu_id = int(sheet2.cell_value(j,3))
    dict1[mänginu_id] = ID_RP[int(sheet2.cell_value(j,3))][0:3]
    dict1[mänginu_id].insert(3,0) #lisan võitude ja kaotuste hinna vahe elemendi
    dict1[mänginu_id].insert(4,0) #lisan võitude ja kaotuste hinna summa elemendi
    dict1[mänginu_id].insert(5,0) #lisan võitude hindade summa elemendi
    dict1[mänginu_id].insert(6,ID_RP[mänginu_id][5])

def master(dict1, sheet2):
    
    rea_counter = 8
    
    for i in dict1:
        hinnacounter = 0 #võitude ja kaotuste hindade jooksev väärtus
        
        RP = dict1[i][1] #reitingupunktid
        
        tulba_counter = 0
        while True:
            vastaseID = sheet2.cell_value(rea_counter, 4 + tulba_counter)
            tulba_counter += 1
            
            if vastaseID == "":
                break
            vastaseID = vastaseID.split(",") #et saada kätte meile oluline vastase ID
            vastaseID = int(vastaseID[0])
            
            if dict1[i][6] != dict1[vastaseID][6]: #mängu ei loeta, kui võistlejad ei ole samast soost
                continue
            
            Reitinguvahe = int(RP - dict1[vastaseID][1])
            
            if 0 <= Reitinguvahe <= 2:
                hinnacounter = 2
            elif 3 <= Reitinguvahe <= 13:
                hinnacounter = 1
            elif Reitinguvahe < 0:
                hinnacounter = (Reitinguvahe + 5)/3
                
            dict1[i][3] += hinnacounter 
            dict1[vastaseID][3] -= hinnacounter
            
            dict1[i][4] += abs(hinnacounter)
            dict1[vastaseID][4] += abs(hinnacounter)
            
            if hinnacounter > 0:
                dict1[i][5] += hinnacounter
            
            
        rea_counter += 1
        
    
    
    
    
    for i in dict1: #reitingu muutuse arvutamine
        hindadesumma = dict1[i][3]
        võiduhindadesumma = dict1[i][5]
        hindadeabs = dict1[i][4]
        kaalud = dict1[i][2]
        
        Reitingumuutus =  round((hindadesumma * 10 + võiduhindadesumma)/(kaalud + hindadeabs))
        
        Kaalumuutus = int(hindadeabs)
        
        UusRP = dict1[i][1] + Reitingumuutus
        Uuskaal = kaalud + Kaalumuutus
        
        PPmuutus = (UusRP - (3 * Uuskaal))/6
        if PPmuutus > 0:
            UusPP = UusRP - PPmuutus
        else:
            UusPP = UusRP
            
        dict1[i][0] = int(UusPP)
        dict1[i][1] = int(UusRP)
        dict1[i][2] = int(Uuskaal)
        dict1[i] = dict1[i][0:3]
        
        
        ID_RP[i][0] = int(UusPP)
        ID_RP[i][1] = int(UusRP)
        ID_RP[i][2] = int(Uuskaal)
        
    return ID_RP

#master(dict1, sheet2)

workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Uusreiting3") #esimese rea kirjutamine
sheet.write(0,0,"personid") 
sheet.write(0,1,"famname")
sheet.write(0,2,"firstname")
sheet.write(0,3,"sex")
sheet.write(0,4,"birthdate")
sheet.write(0,5,"ratedate")
sheet.write(0,6,"rateorder")
sheet.write(0,7,"rateplpnts")
sheet.write(0,8,"ratepoints")
sheet.write(0,9,"rateweight")
sheet.write(0,10,"clbname")

counter = 1
for i in ID_RP:
    sheet.write(counter,0,i)
    sheet.write(counter,1,ID_RP[i][3])
    sheet.write(counter,2,ID_RP[i][4])
    sheet.write(counter,3,ID_RP[i][5])
    sheet.write(counter,4,ID_RP[i][6])
    sheet.write(counter,5,"11/1/2021")
    sheet.write(counter,6,counter)
    sheet.write(counter,7,ID_RP[i][0])
    sheet.write(counter,8,ID_RP[i][1])
    sheet.write(counter,9,ID_RP[i][2])
    sheet.write(counter,10,ID_RP[i][8])
    
    counter += 1

workbook.save("UusreitingProto.xls") #Salvestab uue reitingu soovitud nimega

def kuulõpp(): #Teostab kuulõpus valemi järgi kaalude muutmise
    loc3 = ("UusreitingProto.xls")
    wb3 = xlrd.open_workbook(loc3)
    sheet3 = wb3.sheet_by_index(0)
    
    workbook3 = xlwt.Workbook()
    sheet4 = workbook3.add_sheet("Uusreiting4")
    sheet4.write(0,0,"personid") 
    sheet4.write(0,1,"famname")
    sheet4.write(0,2,"firstname")
    sheet4.write(0,3,"sex")
    sheet4.write(0,4,"birthdate")
    sheet4.write(0,5,"ratedate")
    sheet4.write(0,6,"rateorder")
    sheet4.write(0,7,"rateplpnts")
    sheet4.write(0,8,"ratepoints")
    sheet4.write(0,9,"rateweight")
    sheet4.write(0,10,"clbname")
    
    
    for m in range(1,sheet3.nrows):
        for n in range(11): #kui tegemist on kaaludega, siis rakendatakse valemit
            if n == 9:
                kaal = int(sheet3.cell_value(m,n))
                if kaal > 10:
                    kaal = round(kaal - (kaal**2)/225)
                else:
                    kaal = kaal - 1
                    
                sheet4.write(m,n,kaal)
            else:
                sheet4.write(m,n,sheet3.cell_value(m,n))
                
    workbook3.save("Kuulõpureiting.xls") #uue faili nimi
    
kuulõpp()