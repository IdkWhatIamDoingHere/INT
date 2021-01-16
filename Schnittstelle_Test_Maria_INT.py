import pandas as pd
import openpyxl
from lxml import etree

BWise = "INT_RKM_per_process.xlsx"
Adonis = "2020-12-15_ADONIS-BWise Export_für INT.xlsx"
Users = "2020-12-09_List_of_all_Users.xlsx"
AlleKontrollen = "INT_Schlüsselkontrollen.xlsx"

AdonisImporta = pd.read_excel(Adonis, skiprows = 1)
UsersBWise = pd.read_excel(Users, skiprows = 0)
UsersBWise = UsersBWise[UsersBWise["Name of the users"].str.contains("DE.", na=False)]
UsersBWise = UsersBWise.drop(UsersBWise[UsersBWise["Name of the users"].str.contains("Z_A", na=False)].index)
UsersBWise = UsersBWise["Name of the users"]
UsersBWiseVergleich = []
for user in UsersBWise:
    usermodified = user.lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    UsersBWiseVergleich.append([user,usermodified])
#DataFrame (eingelesene Excel Datei) Filterung

AdonisImportb = AdonisImporta[AdonisImporta["Geschäftsprozessdiagramm - Status"] == "Freigegeben"]
AdonisImportb.insert(0,"Prozess Risiko Code","Prozess Risiko Code")

AdonisImportc = AdonisImportb[AdonisImportb["Risko - Schlüsselrisiko"] == "Ja"]
AdonisImport = AdonisImportc[AdonisImportc["Kontrolle - Schlüsselkontrolle"] == "Ja"]
#Entfernen nicht benötigter Spalten

AdonisImport = AdonisImport.drop(columns = ["Geschäftsprozessdiagramm - Version","Risko - Fehlerhäufigkeit", "Risko - Auswirkung", "Kontrollgruppe - Kontrollnummer", "Kontrollgruppe - Name", "Kontrollgruppe - Kontrollmethode", "Kontrollgruppe - Kontrollausführung", "Kontrollgruppe - Kontrollfrequenz"])

#Übersetzen der einzelnen Zellen gemäß der Übersetzungstabelle
ÜbersetzungRisikotyp = {"Compliance Risiken - 540":"540. Compliance",
    "Financial Reporting Risiken - 660":"660. Abschluss, Hochrechnung, Planung, Financial Statements, Forecast, Planning",
    "Operationelle Risiken - 511 IT Governance":"511. IT Governance",
    "Operationelle Risiken - 512 IT Architektur - IT":"512. Architektur - IT Architecture",
    "Operationelle Risiken - 513 IT Betrieb":"513. IT Betrieb - IT Operations",
    "Operationelle Risiken - 514 Cyber Security":"514. Cyber Security",
    "Operationelle Risiken - 521 Fähigkeiten/Kapazitäten":"521. Fähigkeiten/Kapazitäten.Skills/Capacities",
    "Operationelle Risiken - 522 Verfügbarkeit von Wissen":"522. Verfügbarkeit von Wissen - Availability of Knowledge",
    "Operationelle Risiken - 523 Anreizsysteme":"523. Anreizsysteme - Incentive Systems",
    "Operationelle Risiken - 531 Verträge":"531. Verträge - Contracts",
    "Operationelle Risiken - 532 Haftung und Prozesse":"532. Haftung und Prozesse - Liability and Litigations",
    "Operationelle Risiken - 533 Steuern":"533. Steuern - Tax",
    "Operationelle Risiken - 551 Prozessrisiken":"551. Prozessrisiken - Process Risks",
    "Operationelle Risiken - 552 Projektrisiken":"552 Projektrisiken - Project Risks",
    "Operationelle Risiken - 553 Outscourcing Risiken":"553. In- und Outscourcing Risiken - In- and Outscourcing Risks",
    "Operationelle Risiken - 561 Risikoanalyse":"561. Risikoanalyse und - beweertung - Risk Analysis and Risk Reporting",
    "Operationelle Risiken - 562 Risikoberichterstattung":"562- Risiko-Berichterstattung - Risk Reporting",
    "Operationelle Risiken - 640 Merger & Acquisitions":"640. Merger & Acquisitions",
    "Operationelle Risiken - 651 Externe Kommunikation":"651. Externe Berichterstattung - External Reporting",
    "Operationelle Risiken - 652 Reputationsrisiken":"652. Reputationsmanagement - Reputation Management",
    "Operationelle Risiken - 670 Projektportfolio":"670. Projektportfolio - Project Portfolio",
    "Operationelle Risiken - 680 Interne Fehlinformationen":"680. Interne Fehlinformationen - Internal Misinformation",
    "Kreditrisiken":"220. Kreditrisiken - Credit Risks",
    "Marktpreis- und Liquiditätsrisiken - 220":"21. Marktrisiken - Market Risks"}
AdonisImport["Risko - Fehlerhäufigkeit (technischer Wert)"].replace({"v0":"fast nie (mehr als 5 Jahre)","v1":"selten (alle 5 Jahre)","v2":"gelegentlich (jedes Jahr)","v3":"haeufig (jeden Monat)","v4":"regelmaessig (jede Woche)","v5":"n.a."}, inplace = True)
AdonisImport["Risko - Auswirkung (technischer Wert)"].replace({"v0":"n.a.","v1":"100 Tsd - 1 Mio","v2":"1-3 Mio","v3":"3-5 Mio","v4":"5-10 Mio","v5":"> 10 Mio"}, inplace = True)
AdonisImport["Kontrollgruppe - Kontrollmethode (technischer Wert)"].replace({"v0":"Manuell","v1":"Halbautomatisch","v2":"Automatisch","v3":"<--Kein(e)-->"}, inplace = True)
AdonisImport["Kontrollgruppe - Kontrollausführung (technischer Wert)"].replace({"v0":"Praeventiv","v1":"Detektiv","v2":"<--Kein(e)-->"}, inplace = True)
AdonisImport["Kontrollgruppe - Kontrollfrequenz (technischer Wert)"].replace({"v0":"bei Bedarf","v1":"taeglich","v2":"woechentlich","v3":"monatlich","v4":"vierteljaehrlich","v5":"jaehrlich","v6":"ereignisgesteuert","v7":"<--Kein(e)-->","v9":"halbjaehrlich"}, inplace = True)
AdonisImport["Riskogruppe - Risikotyp"].replace(ÜbersetzungRisikotyp, inplace = True)

i = 0
while i < (len(AdonisImport["Geschäftsprozessdiagramm - Modellname"])) : 
    #AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Modellname"] = "DE.X.XX.XXXX." + AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Modellname"]
    AdonisImport.iloc[i]["Risko - Name"] = AdonisImport.iloc[i]["Risko - Name"].split("_")[(len(AdonisImport.iloc[i]["Risko - Name"].split("_"))-1)]
    AdonisImport.iloc[i]["Kontrolle - Name"] =  AdonisImport.iloc[i]["Kontrolle - Name"][3:]

    
    #if type(AdonisImport.iloc[i]["Kontrolle - Name"]) == float:
        #print("hello") asdjkhlkhdjkfhkjdshfkjhksdhfkjhkjlsdhkljfhkjsdhkljfhkjlshdjkl
    #else:
       #AdonisImport.iloc[i]["Kontrolle - Name"] = AdonisImport.iloc[i]["Kontrolle - Name"].split("_")[(len(AdonisImport.iloc[i]["Kontrolle - Name"].split("_"))-1)]
    
    #print(type(AdonisImport.iloc[i]["Kontrolle - Verantwortlich für Kontrolle"]))
    if type(AdonisImport.iloc[i]["Kontrolle - Verantwortlich für Kontrolle"]) == float:
        
        AdonisImport.iloc[i]["Kontrolle - Verantwortlich für Kontrolle"] = "DE.Sabine Seipold"
        
    else:
    
        AdonisImport.iloc[i]["Kontrolle - Verantwortlich für Kontrolle"] = "DE." + AdonisImport.iloc[i]["Kontrolle - Verantwortlich für Kontrolle"]
   
    if type(AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"]) == float:
        
        AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"] = "DE.Sabine Seipold"
    else:
        
        if "@" in AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"]:
            
            if "Dr." in AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"]:
                
                AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"] = "DE." + AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"].split(".")[0].capitalize() + "." + AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"].split(".")[1].capitalize() + " " + AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"].split(".")[2].split("@")[0].capitalize()
                
            else:
    
                AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"] = "DE." + AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"].split(".")[0].capitalize() + " " + AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"].split(".")[1].split("@")[0].capitalize()
        else:
            AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"] = "DE." + AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"].split(" ")[0].capitalize() + " " + AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Prozesseigner"].split(" ")[1].capitalize()
    
    AdonisImport.iloc[i]["Prozess Risiko Code"] = str(AdonisImport.iloc[i]["Riskogruppe - Risikonummer"]) + "/" + AdonisImport.iloc[i]["Geschäftsprozessdiagramm - Modellname"] + "/" + AdonisImport.iloc[i]["Aufgabe - Name"]
    
    i = i + 1

AdonisImport.rename(columns = {"Geschäftsprozessdiagramm - Modellname":"Prozess Name", "Geschäftsprozessdiagramm - Prozesseigner":"Prozess Owner", "Risko - Name":"Prozess Risiko (PR) Name","Risko - Fehlerhäufigkeit (technischer Wert)":"Eintrittswahrscheinlichkeit", "Risko - Auswirkung (technischer Wert)":"Schadenshoehe", "Schlüsselrisiko":"Schlüsselrisiko?", "Riskogruppe - Beschreibung":"Prozess Risiko Beschreibung", "Kontrolle - Name":"Kontrolle Name" ,"Kontrolle - Handlung im Falle einer Abweichung":"Eskalationsprozess", "Kontrolle - Nachweis der Kontrolle":"Nachweis der durchgeführten Kontrolle", "Kontrolle - Verantwortlich für Kontrolle":"Kontroll Owner","Kontrollgruppe - Beschreibung":"Kontroll Beschreibung", "Kontrollgruppe - Kontrollmethode (technischer Wert)":"Kontrollart", "Kontrollgruppe - Kontrollausführung (technischer Wert)":"Kontrolltyp", "Kontrollgruppe - Kontrollfrequenz (technischer Wert)":"Kontroll-frequenz"}, inplace = True)
AdonisImport.insert(0,"C Validierer",AdonisImport["Prozess Owner"])
AdonisImport.reset_index(drop = True, inplace=True)
############################################################################
NeuerBWiseUser = []
UsersAdonis = AdonisImport["Prozess Owner"]
UsersAdonisZwei = AdonisImport["Kontroll Owner"]
UsersAdonis = UsersAdonis.append(UsersAdonisZwei, ignore_index = True)
UsersAdonis = UsersAdonis.drop_duplicates()
UsersAdonisVergleich = []
for user in UsersAdonis:
    usermodified = user.lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    UsersAdonisVergleich.append([user, usermodified])
    
for useradonis in UsersAdonisVergleich:
    treffer = []
    for userbwise in UsersBWiseVergleich:
        if userbwise[1] == useradonis[1]:
            treffer.append(userbwise)
    if len(treffer) == 0:
        NeuerBWiseUser.append(useradonis)
        a = useradonis[0]
        b = "DE.Sabine Seipold"
        
        AdonisImport.replace({a:b}, inplace = True)
    elif len(treffer) == 1:
        #count = 0
        #for useradoniszwei in UsersAdonisVergleich:
            #if useradoniszwei[1] == treffer[0][1]:
                #count += 1
        #if count == 1:
        a = useradonis[0]
        b = treffer[0][0]
        AdonisImport.replace({a:b}, inplace = True)
            
############################################################################

BWiseAbgleich = pd.read_excel(BWise, skiprows = 3, sheet_name=1)
BWiseAbgleich = BWiseAbgleich.drop(columns = ["Prozesshierarchie  Name", "Prozess Code", "Prozess Legal Entity", "Dokumente", "Kontrolle Code", "Kontrolle legal entity", "Relevante IT Komponente Concat"])
BWiseAbgleich = BWiseAbgleich.drop(BWiseAbgleich[BWiseAbgleich["Prozess Risiko (PR) Name"].str.contains("Fehler in der EUCA", na=False)].index)
BWiseAbgleichA = BWiseAbgleich#[BWiseAbgleich["Schlüsselrisiko?"] == "JA"]


BWiseProzesse = BWiseAbgleichA["Prozess Name"].drop_duplicates()
BWiseProzesseNeu = []
i = 0
for prozess in BWiseProzesse:
    
    #splitprozess = prozess.split(".")[-1].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    splitprozess = prozess[13:].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    
    BWiseProzesseNeu.append([prozess,splitprozess])
    i = i + 1

BWiseRisikoCode = BWiseAbgleichA[["Prozess Risiko Code","Prozess Risiko (PR) Name"]]
BWiseRisikoCode = BWiseRisikoCode.drop_duplicates()

BWiseRisikenNeu = []

for index, risiko in BWiseRisikoCode.iterrows():
    
    if type(risiko["Prozess Risiko Code"]) == float:
        
        continue
    
    #splitprozess = prozess.split(".")[-1].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    splitprozess = risiko["Prozess Risiko Code"].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    
    BWiseRisikenNeu.append([risiko["Prozess Risiko Code"],splitprozess,risiko["Prozess Risiko (PR) Name"]])

BWiseKontrollen = BWiseAbgleichA["Kontrolle Name"].drop_duplicates()
BWiseKontrollenNeu = []
i = 0
for kontrolle in BWiseKontrollen:
    
    if type(kontrolle) == float:
        
        continue
    
    splitkontrolle = kontrolle[11:].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    
    BWiseKontrollenNeu.append([kontrolle,splitkontrolle])
    i = i + 1

########################################################################
    
AdonisProzesseEinzel = AdonisImport["Prozess Name"].drop_duplicates()
AdonisProzesse = []
for prozess in AdonisProzesseEinzel:
    
    #splitprozess = prozess.split(".")[-1].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    splitprozess = prozess.lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    
    AdonisProzesse.append([prozess,splitprozess])

AdonisRisikoCode = AdonisImport[["Prozess Risiko Code", "Prozess Risiko (PR) Name"]]
AdonisRisikoCode = AdonisRisikoCode.drop_duplicates()
AdonisRisikenNeu = []
for index, risiko in AdonisRisikoCode.iterrows():
    
    #splitprozess = prozess.split(".")[-1].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    splitprozess = risiko["Prozess Risiko Code"].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    
    AdonisRisikenNeu.append([risiko["Prozess Risiko Code"],splitprozess, risiko["Prozess Risiko (PR) Name"]])
    

AdonisKontrollen = AdonisImport["Kontrolle Name"].drop_duplicates()
AdonisKontrollenNeu = []
i = 0
for kontrolle in AdonisKontrollen:
    
    splitkontrolle = kontrolle.split("_")[-1].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")
    splitkontrolleNeu = kontrolle.split("_")[-1].strip()
    AdonisKontrollenNeu.append([kontrolle,splitkontrolle, splitkontrolleNeu])
    i = i + 1
    
#######################################################################
NeueProzesseAusAdonis = []
EindeutigeProzesseAusAdonis = []
EindeutigeProzesseAusAdonisUndBWise = []
UneindeutigeProzesseAusAdonis = []

for adonisprozess in AdonisProzesse:
    treffer = []
    for bwiseprozess in BWiseProzesseNeu:
        if bwiseprozess[1] == adonisprozess[1]:
            treffer.append(bwiseprozess)
    if len(treffer) == 0:
        NeueProzesseAusAdonis.append(["DE.X.XX.XXXX."+adonisprozess[0],adonisprozess[0]])
    elif len(treffer) == 1:
        EindeutigeProzesseAusAdonis.append(adonisprozess)
        count = 0
        for adonisprozesszwei in AdonisProzesse:
            if adonisprozesszwei[1] == treffer[0][1]:
                count += 1
        if count == 1:
            EindeutigeProzesseAusAdonisUndBWise.append(treffer[0])
            a = adonisprozess[0]
            b = treffer[0][0]
            
            AdonisImport["Prozess Name"].replace({a:b}, inplace = True)
    else:
        UneindeutigeProzesseAusAdonis.append(adonisprozess)

AlteProzesseAusBWise = []
UneindeutigeProzesseAusBWise = []

for bwiseprozess in BWiseProzesseNeu:
    count = 0
    for adonisprozess in AdonisProzesse:
        if bwiseprozess[1] == adonisprozess[1]:
            count += 1
    if count == 0:
        AlteProzesseAusBWise.append(bwiseprozess)
    elif count == 1:
        continue
    else:
        UneindeutigeProzesseAusBWise.append(bwiseprozess)
##########################################################################
# Risiken
#######################################################################
NeueRisikenAusAdonis = []
EindeutigeRisikenAusAdonis = []
EindeutigeRisikenAusAdonisUndBWise = []
UneindeutigeRisikenAusAdonis = []
risiko_name_num = 1
for adonisrisiko in AdonisRisikenNeu:
    treffer = []
    for bwiserisiko in BWiseRisikenNeu:
        if bwiserisiko[1] == adonisrisiko[1]:
            treffer.append(bwiserisiko)
    if len(treffer) == 0:
        NeueRisikenAusAdonis.append([adonisrisiko[0], "DE.PR.XXX" + str(risiko_name_num) + "." + adonisrisiko[2]])
        
        a = AdonisImport[AdonisImport["Prozess Risiko Code"] == adonisrisiko[0]].index.tolist()
        b = "DE.PR.XXX" + str(risiko_name_num) + "." + adonisrisiko[2]
        
        for index in a:
            #print(str(index) +"     " + AdonisImport.iloc[index]["Prozess Risiko (PR) Name"])
            AdonisImport.iloc[index]["Prozess Risiko (PR) Name"] = b
            
        risiko_name_num += 1
    elif len(treffer) == 1:
        EindeutigeRisikenAusAdonis.append(adonisrisiko)
        count = 0
        for adonisrisikozwei in AdonisRisikenNeu:
            if adonisrisikozwei[1] == treffer[0][1]:
                count += 1
        if count == 1:
            EindeutigeRisikenAusAdonisUndBWise.append(treffer[0])
            a = adonisrisiko[0]
            b = treffer[0][0]
            
            AdonisImport["Prozess Risiko Code"].replace({a:b}, inplace = True)
            
            for index, row in AdonisImport.iterrows():
                if row["Prozess Risiko Code"] == treffer[0][0]:
                    row["Prozess Risiko (PR) Name"] = treffer[0][2]
            
    else:
        print(treffer[0] == treffer[1])
        print(treffer[0])
        print(treffer[1])
        UneindeutigeRisikenAusAdonis.append(adonisrisiko)

AlteRisikenAusBWise = []
UneindeutigeRisikenAusBWise = []

for bwiserisiko in BWiseRisikenNeu:
    
    count = 0
    for adonisrisiko in AdonisRisikenNeu:
        if bwiserisiko[1] == adonisrisiko[1]:
            count += 1  
    if count == 0:
        AlteRisikenAusBWise.append(bwiserisiko)
    elif count == 1:
        continue
    else:
        UneindeutigeRisikenAusBWise.append(bwiserisiko)
##########################################################################
##########################################################################
# KONTROLLEN
#######################################################################
NeueKontrollenAusAdonis = []
EindeutigeKontrollenAusAdonis = []
EindeutigeKontrollenAusAdonisUndBWise = []
UneindeutigeKontrollenAusAdonis = []

for adoniskontrolle in AdonisKontrollenNeu:
    treffer = []
    for bwisekontrolle in BWiseKontrollenNeu:
        if bwisekontrolle[1] == adoniskontrolle[1]:
            treffer.append(bwisekontrolle)
    if len(treffer) == 0:
        NeueKontrollenAusAdonis.append(adoniskontrolle)
    elif len(treffer) == 1:
        EindeutigeKontrollenAusAdonis.append(adoniskontrolle)
        count = 0
        for adoniskontrollezwei in AdonisKontrollenNeu:
            if adoniskontrollezwei[1] == treffer[0][1]:
                count += 1
        if count == 1:
            EindeutigeKontrollenAusAdonisUndBWise.append(treffer[0])
            a = adoniskontrolle[0]
            b = treffer[0][0]
            
            AdonisImport["Kontrolle Name"].replace({a:b}, inplace = True)
    else:
        UneindeutigeKontrollenAusAdonis.append(adoniskontrolle)

AlteKontrollenAusBWise = []
UneindeutigeKontrollenAusBWise = []

for bwisekontrolle in BWiseKontrollenNeu:
    count = 0
    for adoniskontrolle in AdonisKontrollenNeu:
        if bwisekontrolle[1] == adoniskontrolle[1]:
            count += 1
    if count == 0:
        AlteKontrollenAusBWise.append(bwisekontrolle)
    elif count == 1:
        continue
    else:
        UneindeutigeKontrollenAusBWise.append(bwisekontrolle)

###########################################################################
#Liste aller Kontrollen
ListeAllerKontrollen = pd.read_excel(AlleKontrollen, skiprows = 3, sheet_name = 0)
ListeAllerKontrollen = ListeAllerKontrollen["C Name"]
ListeZusätzlicherKontrollen = set(ListeAllerKontrollen) - set(BWiseKontrollen)


ListeAllerKontrollenVergleich = []

for kontrolle in ListeZusätzlicherKontrollen:
    
    if type(kontrolle) == float:
        
        continue
    
    splitkontrolle = kontrolle[11:].lower().strip().replace("ä","ae").replace("ü","ue").replace("ö","oe").replace("ß","ss")  
    ListeAllerKontrollenVergleich.append([kontrolle,splitkontrolle])
    
    
    
NeueKontrollenAusAdonisNeu = []
 
for kontrolleadonis in NeueKontrollenAusAdonis:
    treffer = []
    for kontrollebwise in ListeAllerKontrollenVergleich:
        if kontrollebwise[1] == kontrolleadonis[1]:
            treffer.append(kontrollebwise)
    if len(treffer) == 0:
        NeueKontrollenAusAdonisNeu.append(["DE.PC.XXXX." + kontrolleadonis[2], kontrolleadonis[0], kontrolleadonis[1]])
        
        a = AdonisImport[AdonisImport["Kontrolle Name"] == kontrolleadonis[0]].index.tolist()
        b = "DE.PC.XXXX." + kontrolleadonis[2]
        
        for index in a:
            #print(str(index) +"     " + AdonisImport.iloc[index]["Prozess Risiko (PR) Name"])
            AdonisImport.iloc[index]["Kontrolle Name"] = b
    elif len(treffer) == 1:
        count = 0
        for kontrollenadoniszwei in NeueKontrollenAusAdonis:
            if kontrollenadoniszwei[1] == treffer[0][1]:
                count += 1
        if count == 1:
            a = kontrolleadonis[0]
            b = kontrolleadonis[0][0]
                
            AdonisImport.replace({a:b}, inplace = True)
    else:
        UneindeutigeKontrollenAusAdonis.append(kontrolleadonis)
            
for bwisekontrolle in ListeAllerKontrollenVergleich:
    count = 0
    for adoniskontrolle in NeueKontrollenAusAdonis:
        if bwisekontrolle[1] == adoniskontrolle[1]:
            count += 1
    if count == 0:
        AlteKontrollenAusBWise.append(bwisekontrolle)
    elif count == 1:
        continue
    else:
        UneindeutigeKontrollenAusBWise.append(bwisekontrolle)
        
##########################################################################
# Initalisierung XML Datei

bwise_header = etree.Element("bwiseData", language="de-DE", source="BWise", version="2.0")

bwise_elements = etree.SubElement(bwise_header, "elements")

bwise_links = etree.SubElement(bwise_header, "links")

item_num = 1

##########################################################################
#XML neue Prozesse

VergleichsparameterProzese = ["Prozess Name", "Prozess Owner", "Prozess Risiko (PR) Name"]


if len(NeueProzesseAusAdonis) > 0:
    for neuerprozess in NeueProzesseAusAdonis:
    
        KandidatAdonis = AdonisImport[VergleichsparameterProzese][AdonisImport["Prozess Name"] == neuerprozess[1]]
        KandidatAdonisRelation = KandidatAdonis["Prozess Risiko (PR) Name"]
        ProzessRisikoRelation = []
        for relation in KandidatAdonisRelation:
            ProzessRisikoRelation.append(relation)
        ProzessRisikoRelation = list(set(ProzessRisikoRelation))
        KandidatAdonis = KandidatAdonis.drop(columns=["Prozess Risiko (PR) Name"]).drop_duplicates().reset_index(drop = True).set_index("Prozess Name")
        
        item_id = "xmlid" + str(item_num)
        xml_P_name = str(neuerprozess[0])
        xml_P_owner = str(KandidatAdonis["Prozess Owner"].values[0])
        xml_P_hierarchy = "1.1.Compliance-Management"
    
        bwise_element = etree.SubElement(bwise_elements, "element", type="BusinessProcess", id=item_id)

        etree.SubElement(etree.SubElement(bwise_element, "query", property="label"), "string").text = xml_P_name

        bwise_values = etree.SubElement(bwise_element, "values")

        etree.SubElement(etree.SubElement(bwise_values, "value", property="label"), "string").text = xml_P_name

        etree.SubElement(etree.SubElement(bwise_values, "value", property="description"), "text").text = etree.CDATA("Siehe Prozessdokumentation in Adonis; Veröffentlichung über das Prozesshaus im Intranet der Basler Versicherung Deutschland")

        etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="parentObject"), "refByQuery", type="ProcessHierarchy", property="label"), "string").text = xml_P_hierarchy

        etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="owner"), "refByQuery", type="Resource", property="label"), "string").text = xml_P_owner

        bwise_link = etree.SubElement(bwise_links, "link", association="BusinessProcessRelation")

        bwise_participants = etree.SubElement(bwise_link, "participants")

        etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="relation"), "refByQuery", type="ProcessHierarchy", property="label"), "string").text = xml_P_hierarchy

        etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="context"), "refByQuery", type="Group", property="label"), "string").text = "DE.Basler Deutschland"

        etree.SubElement(etree.SubElement(bwise_participants, "participant", role="subject"), "refById", id=item_id)
        
        for link in ProzessRisikoRelation:
            bwise_link = etree.SubElement(bwise_links, "link", association="RiskRelation")

            bwise_participants = etree.SubElement(bwise_link, "participants")

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="relation"), "refByQuery", type="BusinessProcess", property="label"), "string").text = xml_P_name

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="context"), "refByQuery", type="Group", property="label"), "string").text = "DE.Basler Deutschland"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="subject"), "refByQuery", type="Risk", property="label"), "string").text = link
        item_num += 1  
# XML neue Risiken

VergleichsparameterRisiken = ["Prozess Risiko (PR) Name", "Prozess Risiko Code", "Prozess Risiko Beschreibung", "Eintrittswahrscheinlichkeit", "Schadenshoehe", "Riskogruppe - Risikotyp", "Kontrolle Name"]

if len(NeueRisikenAusAdonis) > 0:
    
    for neuerisiken in NeueRisikenAusAdonis:
            
            KandidatAdonis = AdonisImport[VergleichsparameterRisiken][AdonisImport["Prozess Risiko Code"] == neuerisiken[0]]
            KandidatAdonisRelation = KandidatAdonis["Kontrolle Name"]
            RisikoKontrolleAdonisRelation = []
            for relation in KandidatAdonisRelation:
                RisikoKontrolleAdonisRelation.append(relation)
            RisikoKontrolleAdonisRelation = list(set(RisikoKontrolleAdonisRelation))
            KandidatAdonis = KandidatAdonis.drop(columns=["Kontrolle Name"]).drop_duplicates().reset_index(drop = True).set_index("Prozess Risiko Code")
            
            
            item_id = "xmlid" + str(item_num)
            xml_R_code = str(neuerisiken[0])
            xml_R_name = str(KandidatAdonis["Prozess Risiko (PR) Name"].values[0])
            xml_R_beschreibung = str(KandidatAdonis["Prozess Risiko Beschreibung"].values[0])
            xml_R_owner = "DE.Sabine Seipold"
            xml_R_eintrittswahrscheinlichkeit = str(KandidatAdonis["Eintrittswahrscheinlichkeit"].values[0])
            xml_R_schadenshoehe = str(KandidatAdonis["Schadenshoehe"].values[0])
            xml_R_risikotyp = str(KandidatAdonis["Riskogruppe - Risikotyp"].values[0])
            
            bwise_element = etree.SubElement(bwise_elements, "element", type="Risk", id=item_id)

            # Query
            etree.SubElement(etree.SubElement(bwise_element, "query", property="label"), "string").text = xml_R_name

            # Values

            bwise_values = etree.SubElement(bwise_element, "values")

            # Label
            etree.SubElement(etree.SubElement(bwise_values, "value", property="label"), "string").text = xml_R_name

            # Code
            etree.SubElement(etree.SubElement(bwise_values, "value", property="code"), "string").text = xml_R_code

            # Description
            etree.SubElement(etree.SubElement(bwise_values, "value", property="description"), "text").text = etree.CDATA(xml_R_beschreibung)

            # cust_R_Risikotyp
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_Risikotyp"), "refByQuery", type="CUST_LIST_RISIKOTYP", property="label"), "string").text = "IKS Prozessrisiko"

            # cust_R_ITGC
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_ITGC"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_IKSRELEVANT
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_IKSRELEVANT"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "JA"

            # parentObjekt
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="parentObject"), "refByQuery", type="Risk", property="label"), "string").text = "551.Prozessrisiken - Process Risks"

            # cust_R_OPERATIONELLES
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_OPERATIONELLES"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "JA"

            # cust_R_RMORSARELEVANT
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_RMORSARELEVANT"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_KEYRISK
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_KEYRISK"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "JA"

            # cust_R_OWNER
            etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_OWNER"), "string").text = "DE.Sabine Seipold"

            # cust_R_ELC
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_ELC"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_COMPLIANCE
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_COMPLIANCE"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_FINANZELLE
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_FINANZELLE"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_RISKPROBABILITY
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_RISKPROBABILITY"), "refByQuery", type="CUST_LIST_PROBABILITY", property="label"), "string").text = xml_R_eintrittswahrscheinlichkeit

            # cust_R_RISK
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_RISK"), "refByQuery", type="CUST_LIST_RISKVALUE", property="label"), "string").text = xml_R_schadenshoehe

            # cust_R_Curreny
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_Currency"), "refByQuery", type="CUST_LIST_WAEHRUNG_EUR_CHF", property="label"), "string").text = "EUR"

            # Links

            bwise_link = etree.SubElement(bwise_links, "link", association="RiskDivisionRelation")

            bwise_participants = etree.SubElement(bwise_link, "participants")

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="subject"), "refByQuery", type="Risk", property="label"), "string").text = xml_R_name

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="relation"), "refByQuery", type="Group", property="label"), "string").text = "DE.Basler Deutschland"

            for link in RisikoKontrolleAdonisRelation:
                
                bwise_link = etree.SubElement(bwise_links, "link", association="ControlMeasureRelation")

                bwise_participants = etree.SubElement(bwise_link, "participants")

                etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="relation"), "refByQuery", type="Risk", property="label"), "string").text = link

                etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="context"), "refByQuery", type="Group", property="label"), "string").text = "DE.Basler Deutschland"

                etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="subject"), "refByQuery", type="ControlMeasure", property="label"), "string").text = xml_R_name

            item_num += 1
            
# XML neue Kontrollen

VergleichsparameterKontrolle = ["Kontrolle Name","Eskalationsprozess", "Nachweis der durchgeführten Kontrolle", "Kontroll Owner", "Kontroll Beschreibung", "Kontrollart", "Kontrolltyp", "Kontroll-frequenz", "C Validierer"]

if len(NeueKontrollenAusAdonisNeu) > 0:
    
    for neuekontrollen in NeueKontrollenAusAdonisNeu:
        
            KandidatAdonis = AdonisImport[VergleichsparameterKontrolle][AdonisImport["Kontrolle Name"] == neuekontrollen[0]].drop_duplicates().reset_index(drop = True).set_index("Kontrolle Name")
            
            item_id = "xmlid" + str(item_num)
            xml_K_name = str(neuekontrollen[0])
            xml_K_beschreibung = str(KandidatAdonis["Kontroll Beschreibung"].values[0])
            xml_K_owner = str(KandidatAdonis["Kontroll Owner"].values[0])
            xml_K_eskalationsprozess = str(KandidatAdonis["Eskalationsprozess"].values[0])
            xml_K_kontrollfrequenz = str(KandidatAdonis["Kontroll-frequenz"].values[0])
            xml_K_kontrollart = str(KandidatAdonis["Kontrollart"].values[0])
            xml_K_kontrolltyp = str(KandidatAdonis["Kontrolltyp"].values[0])
            xml_K_validierer = str(KandidatAdonis["C Validierer"].values[0])
            
            bwise_element = etree.SubElement(bwise_elements, "element", type="ControlMeasure", id=item_id)

            bwise_query = etree.SubElement(bwise_element, "query", property="label")
            etree.SubElement(bwise_query, "string").text = xml_K_name

            bwise_values = etree.SubElement(bwise_element, "values")

            bwise_parentObject = etree.SubElement(bwise_values, "value", property="parenObject")
            etree.SubElement(etree.SubElement(bwise_parentObject, "refByQuery", type="ControlMeasureCategory", property="label"), "string").text = "1.3 Prozesskontrollen implementiert"

            bwise_label = etree.SubElement(bwise_values, "value", property="label")
            etree.SubElement(bwise_label, "string").text = xml_K_name

            bwise_owner = etree.SubElement(bwise_values, "value", property="owner")
            etree.SubElement(etree.SubElement(bwise_owner, "refByQuery", type="Resource", property="label"), "string").text = xml_K_owner

            bwise_description = etree.SubElement(bwise_values, "value", property="description")
            etree.SubElement(bwise_description, "text").text = etree.CDATA(xml_K_beschreibung)

            bwise_evidenceDefinition = etree.SubElement(bwise_values, "value", property="evidenceDefinition")
            etree.SubElement(bwise_evidenceDefinition, "text").text = etree.CDATA("noch nicht")

            bwise_cust_C_ESCALATIONPROCESS = etree.SubElement(bwise_values, "value", property="cust_C_ESCALATIONPROCESS")
            etree.SubElement(bwise_cust_C_ESCALATIONPROCESS, "text").text = etree.CDATA(xml_K_eskalationsprozess)

            bwise_cust_C_CONTROLFEQUENCY = etree.SubElement(bwise_values, "value", property="cust_C_CONTROLFEQUENCY")
            etree.SubElement(etree.SubElement(bwise_cust_C_CONTROLFEQUENCY, "refByQuery", type="CUST_LIST_CONTROLFREQUENCY", property="label"), "string").text = xml_K_kontrollfrequenz

            bwise_cust_DEPENDENT_ON_IT = etree.SubElement(bwise_values, "value", property="cust_DEPENDENT_ON_IT")
            etree.SubElement(etree.SubElement(bwise_cust_DEPENDENT_ON_IT, "refByQuery", type="CUST_LIST_ITDEPENDENTTYPE", property="label"), "string").text = xml_K_kontrollart

            bwise_cust_C_KEYCONTROL = etree.SubElement(bwise_values, "value", property="cust_C_KEYCONTROL")
            etree.SubElement(etree.SubElement(bwise_cust_C_KEYCONTROL, "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "JA"

            bwise_cust_CONTROL_TYPE = etree.SubElement(bwise_values, "value", property="cust_CONTROL_TYPE")
            etree.SubElement(etree.SubElement(bwise_cust_CONTROL_TYPE, "refByQuery", type="CUST_LIST_CONTROLTYPE", property="label"), "string").text = xml_K_kontrolltyp
##########################################################################
            #XML Bewertungs-Setup
            item_num += 1
            item_id = "xmlid" + str(item_num)
            bwise_element = etree.SubElement(bwise_elements, "element", type="BWCntrlMonitorAssm", id=item_id)

# Query

            bwise_values = etree.SubElement(bwise_element, "values")

# Label

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="subjects"), "refByQuery", type="ControlMeasure", property="label"), "string").text = xml_K_name

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="reviewFrequency"), "refByQuery", type="ControlSchedule", property="label"), "string").text = "DE: Gutachten"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="group"), "refByQuery", type="Group", property="label"), "string").text = "DE.Basler Deutschland"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="tester"), "refByQuery", type="Resource", property="label"), "string").text ="Eine Person"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="monitorFrequency"), "refByQuery", type="ControlSchedule", property="label"), "string").text = "DE: Validierung"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="monitor"), "refByQuery", type="Resource", property="label"), "string").text = xml_K_validierer

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="reviewer"), "refByQuery", type="Resource", property="label"), "string").text = xml_K_owner
##########################################################################
            item_num += 1

##########################################################################
#Vergleich - Prozesse
##########################################################################
VeränderungenProzesse = None
VergleichsparameterProzese = ["Prozess Name", "Prozess Owner", "Prozess Risiko (PR) Name"]
durchgang = 0

for eindeutigerprozess in EindeutigeProzesseAusAdonisUndBWise:
    
    KandidatBWise = BWiseAbgleichA[VergleichsparameterProzese][BWiseAbgleichA["Prozess Name"] == eindeutigerprozess[0]]
    KandidatBWiseRelation = KandidatBWise["Prozess Risiko (PR) Name"]
    ProzessRisikoBWiseRelation = []
    for relation in KandidatBWiseRelation:
        ProzessRisikoBWiseRelation.append(relation)
    KandidatBWise = KandidatBWise.drop(columns=["Prozess Risiko (PR) Name"]).drop_duplicates().reset_index(drop = True).set_index("Prozess Name")
    ProzessRisikoBWiseRelation = list(set(ProzessRisikoBWiseRelation))
    KandidatBWise["Verknüpfte Risiken"] = [ProzessRisikoBWiseRelation]
    
    KandidatAdonis = AdonisImport[VergleichsparameterProzese][AdonisImport["Prozess Name"] == eindeutigerprozess[0]]
    KandidatAdonisRelation = KandidatAdonis["Prozess Risiko (PR) Name"]
    ProzessRisikoAdonisRelation = []
    for relation in KandidatAdonisRelation:
        ProzessRisikoAdonisRelation.append(relation)
    KandidatAdonis = KandidatAdonis.drop(columns=["Prozess Risiko (PR) Name"]).drop_duplicates().reset_index(drop = True).set_index("Prozess Name")
    ProzessRisikoAdonisRelation = list(set(ProzessRisikoAdonisRelation))
    KandidatAdonis["Verknüpfte Risiken"] = [ProzessRisikoAdonisRelation]
    
    if KandidatBWise.equals(KandidatAdonis) == False:
        
        alte_verknüpfte_risiken = list(set(ProzessRisikoBWiseRelation) - set(ProzessRisikoAdonisRelation))
        neue_verknüpfte_risiken = list(set(ProzessRisikoAdonisRelation) - set(ProzessRisikoBWiseRelation))
        DF_compare_Prozesse = KandidatBWise.compare(KandidatAdonis, keep_shape = True)
        DF_compare_Prozesse["Alter Verknüpfungen"] = [alte_verknüpfte_risiken]
        DF_compare_Prozesse["Neue Verknüpfungen"] = [neue_verknüpfte_risiken]
        
        if durchgang > 0:
            VeränderungenProzesse = VeränderungenProzesse.append(DF_compare_Prozesse, ignore_index = False)
        else:
            VeränderungenProzesse = DF_compare_Prozesse
            durchgang += 1
            
        item_id = "xmlid" + str(item_num)
        xml_P_name = str(KandidatAdonis.index[0])
        xml_P_owner = str(KandidatAdonis["Prozess Owner"].values[0])
    
        bwise_element = etree.SubElement(bwise_elements, "element", type="BusinessProcess", id=item_id)

        etree.SubElement(etree.SubElement(bwise_element, "query", property="label"), "string").text = xml_P_name

        bwise_values = etree.SubElement(bwise_element, "values")

        etree.SubElement(etree.SubElement(bwise_values, "value", property="label"), "string").text = xml_P_name

        etree.SubElement(etree.SubElement(bwise_values, "value", property="description"), "text").text = etree.CDATA("Siehe Prozessdokumentation in Adonis; Veröffentlichung über das Prozesshaus im Intranet der Basler Versicherung Deutschland")

        etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="owner"), "refByQuery", type="Resource", property="label"), "string").text = xml_P_owner
        
        for link in neue_verknüpfte_risiken:
            bwise_link = etree.SubElement(bwise_links, "link", association="RiskRelation")

            bwise_participants = etree.SubElement(bwise_link, "participants")

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="relation"), "refByQuery", type="BusinessProcess", property="label"), "string").text = xml_P_name

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="context"), "refByQuery", type="Group", property="label"), "string").text = "DE.Basler Deutschland"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="subject"), "refByQuery", type="Risk", property="label"), "string").text = link
    
        item_num += 1  
        
##########################################################################
#Vergleich - Risiken
##########################################################################
VeränderungenRisiken =None
VergleichsparameterRisiken = ["Prozess Risiko (PR) Name", "Prozess Risiko Code", "Prozess Risiko Beschreibung", "Eintrittswahrscheinlichkeit", "Schadenshoehe", "Kontrolle Name"]
durchgang = 0

for eindeutigesrisiko in EindeutigeRisikenAusAdonisUndBWise:
    
    KandidatBWise = BWiseAbgleichA[VergleichsparameterRisiken][BWiseAbgleichA["Prozess Risiko Code"] == eindeutigesrisiko[0]]
    
    KandidatBWiseRelation = KandidatBWise["Kontrolle Name"]
    RisikoKontrolleBWiseRelation = []
    for relation in KandidatBWiseRelation:
        RisikoKontrolleBWiseRelation.append(relation)
    KandidatBWise = KandidatBWise.drop(columns=["Kontrolle Name"]).drop_duplicates().reset_index(drop = True).set_index("Prozess Risiko Code")
    RisikoKontrolleBWiseRelation = list(set(RisikoKontrolleBWiseRelation))
    if len(KandidatBWise) == 1:
        KandidatBWise["Verknüpfte Kontrollen"] = [RisikoKontrolleBWiseRelation]
    else:
        print("Achtung")
    
    
    KandidatAdonis = AdonisImport[VergleichsparameterRisiken][AdonisImport["Prozess Risiko Code"] == eindeutigesrisiko[0]]
    
    KandidatAdonisRelation = KandidatAdonis["Kontrolle Name"]
    RisikoKontrolleAdonisRelation = []
    for relation in KandidatAdonisRelation:
        RisikoKontrolleAdonisRelation.append(relation)
    KandidatAdonis = KandidatAdonis.drop(columns=["Kontrolle Name"]).drop_duplicates().reset_index(drop = True).set_index("Prozess Risiko Code")
    RisikoKontrolleAdonisRelation = list(set(RisikoKontrolleAdonisRelation))
    if len(KandidatAdonis) == 1:
        KandidatAdonis["Verknüpfte Kontrollen"] = [RisikoKontrolleAdonisRelation]
    else:
        print("Achtung 2")
    
    
    
    if KandidatBWise.equals(KandidatAdonis) == False:
        
        if len(KandidatAdonis) == 1:
            
            alte_verknüpfte_kontrollen = list(set(RisikoKontrolleBWiseRelation) - set(RisikoKontrolleAdonisRelation))
            neue_verknüpfte_kontrollen = list(set(RisikoKontrolleAdonisRelation) - set(RisikoKontrolleBWiseRelation))
            DF_compare_Risiken = KandidatBWise.compare(KandidatAdonis, keep_shape = True)
            DF_compare_Risiken["Alter Verknüpfungen"] = [alte_verknüpfte_kontrollen]
            DF_compare_Risiken["Neue Verknüpfungen"] = [neue_verknüpfte_kontrollen]
            
            if durchgang > 0:
                
                if len(KandidatBWise.index) > 1:
                    continue
                else:
                    VeränderungenRisiken = VeränderungenRisiken.append(DF_compare_Risiken)
            else:
                VeränderungenRisiken = DF_compare_Risiken
                durchgang += 1
                
# XML Risiko erstellen
            item_id = "xmlid" + str(item_num)
            xml_R_code = str(KandidatAdonis.index[0])
            xml_R_name = str(KandidatAdonis["Prozess Risiko (PR) Name"].values[0])
            xml_R_beschreibung = str(KandidatAdonis["Prozess Risiko Beschreibung"].values[0])
            xml_R_owner = "DE.Sabine Seipold"
            xml_R_eintrittswahrscheinlichkeit = str(KandidatAdonis["Eintrittswahrscheinlichkeit"].values[0])
            xml_R_schadenshoehe = str(KandidatAdonis["Schadenshoehe"].values[0])
            
            bwise_element = etree.SubElement(bwise_elements, "element", type="Risk", id=item_id)

            # Query
            etree.SubElement(etree.SubElement(bwise_element, "query", property="label"), "string").text = xml_R_name

            # Values

            bwise_values = etree.SubElement(bwise_element, "values")

            # Label
            etree.SubElement(etree.SubElement(bwise_values, "value", property="label"), "string").text = xml_R_name

            # Code
            etree.SubElement(etree.SubElement(bwise_values, "value", property="code"), "string").text = xml_R_code

            # Description
            etree.SubElement(etree.SubElement(bwise_values, "value", property="description"), "text").text = etree.CDATA(xml_R_beschreibung)

            # cust_R_Risikotyp
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_Risikotyp"), "refByQuery", type="CUST_LIST_RISIKOTYP", property="label"), "string").text = "IKS Prozessrisiko"

            # cust_R_ITGC
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_ITGC"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_IKSRELEVANT
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_IKSRELEVANT"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "JA"

            # parentObjekt
            # Mal schauen ob es notwendig ist

            # cust_R_OPERATIONELLES
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_OPERATIONELLES"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "JA"

            # cust_R_RMORSARELEVANT
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_RMORSARELEVANT"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_KEYRISK
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_KEYRISK"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "JA"

            # cust_R_OWNER
            etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_OWNER"), "string").text = "DE.Sabine Seipold"

            # cust_R_ELC
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_ELC"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_COMPLIANCE
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_COMPLIANCE"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_FINANZELLE
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_FINANZELLE"), "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "NEIN"

            # cust_R_RISKPROBABILITY
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_RISKPROBABILITY"), "refByQuery", type="CUST_LIST_PROBABILITY", property="label"), "string").text = xml_R_eintrittswahrscheinlichkeit

            # cust_R_RISK
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_RISK"), "refByQuery", type="CUST_LIST_RISKVALUE", property="label"), "string").text = xml_R_schadenshoehe

            # cust_R_Curreny
            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="cust_R_Currency"), "refByQuery", type="CUST_LIST_WAEHRUNG_EUR_CHF", property="label"), "string").text = "EUR"

            # Links

            bwise_link = etree.SubElement(bwise_links, "link", association="RiskDivisionRelation")

            bwise_participants = etree.SubElement(bwise_link, "participants")

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="subject"), "refByQuery", type="Risk", property="label"), "string").text = xml_R_name

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="relation"), "refByQuery", type="Group", property="label"), "string").text = "DE.Basler Deutschland"
            
            for link in neue_verknüpfte_kontrollen:
                
                bwise_link = etree.SubElement(bwise_links, "link", association="ControlMeasureRelation")

                bwise_participants = etree.SubElement(bwise_link, "participants")

                etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="relation"), "refByQuery", type="Risk", property="label"), "string").text = link

                etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="context"), "refByQuery", type="Group", property="label"), "string").text = "DE.Basler Deutschland"

                etree.SubElement(etree.SubElement(etree.SubElement(bwise_participants, "participant", role="subject"), "refByQuery", type="ControlMeasure", property="label"), "string").text = xml_R_name
            item_num += 1
##########################################################################
#Vergleich - Kontrollen
##########################################################################
VeränderungenKontrollen = None
VergleichsparameterKontrolle = ["Kontrolle Name","Eskalationsprozess", "Nachweis der durchgeführten Kontrolle", "Kontroll Owner", "Kontroll Beschreibung", "Kontrollart", "Kontrolltyp", "Kontroll-frequenz", "C Validierer"]
durchgang = 0

for eindeutigekontrolle in EindeutigeKontrollenAusAdonisUndBWise:
    
    KandidatBWise = BWiseAbgleichA[VergleichsparameterKontrolle][BWiseAbgleichA["Kontrolle Name"] == eindeutigekontrolle[0]].drop_duplicates().reset_index(drop = True).set_index("Kontrolle Name")
    KandidatAdonis = AdonisImport[VergleichsparameterKontrolle][AdonisImport["Kontrolle Name"] == eindeutigekontrolle[0]].drop_duplicates().reset_index(drop = True).set_index("Kontrolle Name")
    
    if KandidatBWise.equals(KandidatAdonis) == False:
        if len(KandidatAdonis) == 1:
            if durchgang > 0:
                VeränderungenKontrollen = VeränderungenKontrollen.append(KandidatBWise.compare(KandidatAdonis, keep_shape = True), ignore_index = False)
                
            else:        
                VeränderungenKontrollen = KandidatBWise.compare(KandidatAdonis, keep_shape = True)
                durchgang += 1

# XML Kontrollen erstellen
            item_id = "xmlid" + str(item_num)
            xml_K_name = str(KandidatAdonis.index[0])
            xml_K_beschreibung = str(KandidatAdonis["Kontroll Beschreibung"].values[0])
            xml_K_owner = str(KandidatAdonis["Kontroll Owner"].values[0])
            xml_K_eskalationsprozess = str(KandidatAdonis["Eskalationsprozess"].values[0])
            xml_K_kontrollfrequenz = str(KandidatAdonis["Kontroll-frequenz"].values[0])
            xml_K_kontrollart = str(KandidatAdonis["Kontrollart"].values[0])
            xml_K_kontrolltyp = str(KandidatAdonis["Kontrolltyp"].values[0])
            xml_K_validierer = str(KandidatAdonis["C Validierer"].values[0])
            
                
            bwise_element = etree.SubElement(bwise_elements, "element", type="ControlMeasure", id=item_id)

            bwise_query = etree.SubElement(bwise_element, "query", property="label")
            etree.SubElement(bwise_query, "string").text = xml_K_name

            bwise_values = etree.SubElement(bwise_element, "values")

            #bwise_parentObject = etree.SubElement(bwise_values, "value", property="parenObject")
            #etree.SubElement(etree.SubElement(bwise_parentObject, "refByQuery", type="ControlMeasureCategory", property="label"), "string").text = "1.3.1 Prozesskontrollen implementiert"

            bwise_label = etree.SubElement(bwise_values, "value", property="label")
            etree.SubElement(bwise_label, "string").text = xml_K_name

            bwise_owner = etree.SubElement(bwise_values, "value", property="owner")
            etree.SubElement(etree.SubElement(bwise_owner, "refByQuery", type="Resource", property="label"), "string").text = xml_K_owner

            bwise_description = etree.SubElement(bwise_values, "value", property="description")
            etree.SubElement(bwise_description, "text").text = etree.CDATA(xml_K_beschreibung)

            bwise_evidenceDefinition = etree.SubElement(bwise_values, "value", property="evidenceDefinition")
            etree.SubElement(bwise_evidenceDefinition, "text").text = etree.CDATA("noch nicht")

            bwise_cust_C_ESCALATIONPROCESS = etree.SubElement(bwise_values, "value", property="cust_C_ESCALATIONPROCESS")
            etree.SubElement(bwise_cust_C_ESCALATIONPROCESS, "text").text = etree.CDATA(xml_K_eskalationsprozess)

            bwise_cust_C_CONTROLFEQUENCY = etree.SubElement(bwise_values, "value", property="cust_C_CONTROLFEQUENCY")
            etree.SubElement(etree.SubElement(bwise_cust_C_CONTROLFEQUENCY, "refByQuery", type="CUST_LIST_CONTROLFREQUENCY", property="label"), "string").text = xml_K_kontrollfrequenz

            bwise_cust_DEPENDENT_ON_IT = etree.SubElement(bwise_values, "value", property="cust_DEPENDENT_ON_IT")
            etree.SubElement(etree.SubElement(bwise_cust_DEPENDENT_ON_IT, "refByQuery", type="CUST_LIST_ITDEPENDENTTYPE", property="label"), "string").text = xml_K_kontrollart

            bwise_cust_C_KEYCONTROL = etree.SubElement(bwise_values, "value", property="cust_C_KEYCONTROL")
            etree.SubElement(etree.SubElement(bwise_cust_C_KEYCONTROL, "refByQuery", type="CUST_LIST_JA_NEIN", property="label"), "string").text = "JA"

            bwise_cust_CONTROL_TYPE = etree.SubElement(bwise_values, "value", property="cust_CONTROL_TYPE")
            etree.SubElement(etree.SubElement(bwise_cust_CONTROL_TYPE, "refByQuery", type="CUST_LIST_CONTROLTYPE", property="label"), "string").text = xml_K_kontrolltyp

            item_num += 1
            item_id = "xmlid" + str(item_num)
            bwise_element = etree.SubElement(bwise_elements, "element", type="BWCntrlMonitorAssm", id=item_id)

# Query

            bwise_values = etree.SubElement(bwise_element, "values")

# Label

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="subjects"), "refByQuery", type="ControlMeasure", property="label"), "string").text = xml_K_name

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="reviewFrequency"), "refByQuery", type="ControlSchedule", property="label"), "string").text = "DE: Gutachten"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="group"), "refByQuery", type="Group", property="label"), "string").text = "DE.Basler Deutschland"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="tester"), "refByQuery", type="Resource", property="label"), "string").text ="Eine Person"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="monitorFrequency"), "refByQuery", type="ControlSchedule", property="label"), "string").text = "DE: Validierung"

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="monitor"), "refByQuery", type="Resource", property="label"), "string").text = xml_K_validierer

            etree.SubElement(etree.SubElement(etree.SubElement(bwise_values, "value", property="reviewer"), "refByQuery", type="Resource", property="label"), "string").text = xml_K_owner
##########################################################################
            
            item_num += 1
            
        else:
            print("Vorsicht!")
            
            
#########################################################################

    
#########################################################################
#Export
#########################################################################

# XML Export
tree = etree.ElementTree(bwise_header)
tree.write("XML_Alles.xml")

#############################################

with pd.ExcelWriter("mit_XML_INT.xlsx") as writer:
    pd.Series(NeueProzesseAusAdonis).to_excel(writer, sheet_name="Neue Prozesse aus Adonis", index=False)
    #pd.Series(uneindeutigeProzesseAB).to_excel(writer, sheet_name="Uneindeutige Prozesse in BWise", index=False)
    pd.Series(AlteProzesseAusBWise).to_excel(writer, sheet_name="Überflüssige Prozesse aus BWise", index=False)
    #pd.Series(uneindeutigeProzesseBA).to_excel(writer, sheet_name="Uneindeutige Prozesse in Adonis", index=False)
    pd.Series(NeueRisikenAusAdonis).to_excel(writer, sheet_name="Neue Risiken aus Adonis", index=False)
    #pd.Series(uneindeutigeRisikenAB).to_excel(writer, sheet_name="Uneindeutige Risiken in BWise", index=False)
    pd.Series(AlteRisikenAusBWise).to_excel(writer, sheet_name="Überflüssige Risiken aus BWise", index=False)
    #pd.Series(uneindeutigeRisikenBA).to_excel(writer, sheet_name="Uneindeutige RIsiken in Adonis ", index=False)
    pd.Series(NeueKontrollenAusAdonisNeu).to_excel(writer, sheet_name="Neue Kontrollen aus Adonis", index=False)
    pd.Series(UneindeutigeKontrollenAusAdonis).to_excel(writer, sheet_name="Uneindeutige Kontrollen in BWise", index=False)
    pd.Series(AlteKontrollenAusBWise).to_excel(writer, sheet_name="Überflüssige Kontrollen aus BWise", index=False)
    pd.Series(UneindeutigeKontrollenAusBWise).to_excel(writer, sheet_name="Uneindeutige Kontrollen in Adonis", index=False)
    if type(VeränderungenProzesse) != str:
       VeränderungenProzesse.to_excel(writer, sheet_name="Änderungen Prozesse")
    if type(VeränderungenRisiken) != str:
        VeränderungenRisiken.to_excel(writer, sheet_name="Änderungen Risiken")
    VeränderungenKontrollen.to_excel(writer, sheet_name="Änderungen Kontrollen")
    pd.Series(NeuerBWiseUser).to_excel(writer, sheet_name="Neue BWise Users", index=False)
##############################################