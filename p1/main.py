import xml.etree.ElementTree as ET
from os import listdir, rename, mkdir
from os.path import isfile, join
import datetime
import Bimail
import calendar
from pathlib import Path

ag_mapping = dict()
tour_mapping = dict()

againLines = []
msgs = []
inputErrorCount = 0

todaySTR = datetime.datetime.today().strftime('%Y-%m-%d')
auftrag = 0
recipients = []

outFolder = ""
inFolder = ""
errFolder = ""
AgMappingPath = ""
TourMappingPath = ""
sourcePath = ""

def readCSV(path):

    global inputErrorCount
    global auftrag

    with open(path, 'r') as f:
        lines = f.readlines()

        for line in lines:
            prnt = True

            tree = ET.parse(sourcePath)
            root = tree.getroot()

            NK = root.find('NK')
            AK = NK.find('AK')
            SD = AK.find('SD')
            TOUR = AK.find('TOURVORGABE')
            RSTK = SD.find('RSTK')

            lineElems = line.split(";")

        #Datums
            dat = lineElems[1]
            if len(lineElems[0]) < 2:
                dat = dat + "0"
            dat = dat + lineElems[0]

            #erster, letzter
            fom = dat + "01"
            lom = dat + str(calendar.monthrange(int(lineElems[1]), int(lineElems[0]))[1])

            nkdatum = NK.find('NKDATUM')
            nkdatum.text = fom

            ladedatum = RSTK.find('LADEDATUM')
            ladedatum.text = fom
            ladedatumbis = RSTK.find('LADEDATUMBIS')
            ladedatumbis.text = fom

            lieferdatum = RSTK.find('LIEFERDATUM')
            lieferdatum.text = lom
            lieferdatumbis = RSTK.find('LIEFERDATUMBIS')
            lieferdatumbis.text = lom

            tourstart = TOUR.find('TOURSTARTDATUM')
            tourstart.text = fom
            tourende = TOUR.find('TOURENDEDATUM')
            tourende.text = lom

        #AG/ABS
            if lineElems[6] not in ag_mapping:
                inputErrorCount = inputErrorCount + 1
                againLines.append(line)
                text = "Zeile: " + str(inputErrorCount) + " - Konnte Auftraggeber " + lineElems[6] + " nicht finden"
                msgs.append(text)
                prnt = False
            else:
                adrag = AK.find('ADR-AG')
                agnr = adrag.find('NR')
                agnr.text = ag_mapping[lineElems[6]]

                adrabs = SD.find('ADR-ABS')
                absnr = adrabs.find('NR')
                absnr.text = ag_mapping[lineElems[6]]

        #EMP/Entladestelle
            #UNGÜLTIG
            if lineElems[5] not in tour_mapping:
                inputErrorCount = inputErrorCount + 1
                againLines.append(line)
                text = "Zeile: " + str(inputErrorCount) + " - Konnte Tour " + lineElems[5] + " nicht finden"
                msgs.append(text)
                prnt = False
            else:
            #GÜLTIG
                adremp = SD.find('ADR-EMP')
                #Z'FUAS
                if tour_mapping[lineElems[5]][1] == '9999':
                    name = ET.Element('NAME')
                    name.text = tour_mapping[lineElems[5]][2]
                    adremp.append(name)
                    ort = ET.Element('ORT1')
                    ort.text = tour_mapping[lineElems[5]][3]
                    adremp.append(ort)
                    land = ET.Element('LAND')
                    land.text = tour_mapping[lineElems[5]][4]
                    adremp.append(land)
                    isoland = ET.Element('ISOLAND')
                    isoland.text = tour_mapping[lineElems[5]][5]
                    adremp.append(isoland)
                    plz = ET.Element('PLZ')
                    plz.text = tour_mapping[lineElems[5]][6]
                    adremp.append(plz)
                    strasse = ET.Element('STRASSE')
                    strasse.text = tour_mapping[lineElems[5]][7]
                    adremp.append(strasse)
                    hausnr = ET.Element('HAUSNR')
                    hausnr.text = tour_mapping[lineElems[5]][8]
                    adremp.append(hausnr)
                #HINTERLEGT
                else:
                    nr = ET.Element('NR')
                    nr.text = tour_mapping[lineElems[5]][1]
                    adremp.append(nr)

        #GEWICHT
            PO = SD.find('PO')
            POFD = PO.find('POFD')
            tgewicht = POFD.find('TGEWICHT')
            tgewicht.text = lineElems[7]
            ugewicht = POFD.find('UGEWICHT')
            ugewicht.text = lineElems[7]
            fgewicht = POFD.find('FGEWICHT')
            fgewicht.text = lineElems[7]

        #EXTNR
            extnr = AK.find('EXTNR')
            extnr.text = "Tour " + lineElems[5]

        #OUTPUT
            if prnt:
                auftrag = auftrag + 1
                tree.write(outFolder + "/AT_" + todaySTR + "_" + str(auftrag) + ".xml")


def loadConfig():

    global recipients
    global inFolder
    global outFolder
    global errFolder
    global AgMappingPath
    global TourMappingPath
    global sourcePath

    with open('config.csv', 'r') as f:
        lines = f.readlines()
        print(lines)
        for x in range(len(lines)):
            lineElems = lines[x].split(';')

            print(x)
            print(lineElems)

            if x == 0:
                for y in range(1, len(lineElems)):
                    recipients.append(lineElems[y])

            if x == 1:
                if lineElems[2] == '1':
                    Path(lineElems[1]).mkdir(parents=True, exist_ok=True)
                inFolder = lineElems[1]


            if x == 2:
                if lineElems[2] == '1':
                    Path(lineElems[1]).mkdir(parents=True, exist_ok=True)
                outFolder = lineElems[1]

            if x == 3:
                if lineElems[2] == '1':
                    Path(lineElems[1]).mkdir(parents=True, exist_ok=True)
                errFolder = lineElems[1]

            if x == 4:
                AgMappingPath = lineElems[1]

            if x == 5:
                TourMappingPath = lineElems[1]

            if x == 6:
                sourcePath = lineElems[1]


def loadAGmapping():
    with open(AgMappingPath, 'r') as f:
        lines = f.readlines()

        for line in lines:
            l = line.split(';')
            ag_mapping[l[0]] = l[1]

def loadTourmapping():
    with open(TourMappingPath, 'r') as f:
        lines = f.readlines()

        for line in lines:
            l = line.split(';')
            # Ganze zeile als value, erster eintrag als key
            tour_mapping[l[0]] = l

def handleErrors():

#Schreibe Neue CSV
    errorFile = errFolder+"/"+todaySTR+".csv"

    with open(errorFile, 'w') as w:
        w.writelines(againLines)
        w.close()

    subj = "Jäger Importe"
    text = "Einige Zeilen konnten nicht importiert werden.<br>Diese Zeilen wurden in der angehängten Datei herausgefiltert - der Rest ist importiert worden.<br><br>Zeile (neue Datei) + Begründung: <br><br>"
    for msg in msgs:
        text = text + msg + "<br>\n"

    mymail = Bimail.Bimail(subj, recipients)
    mymail.htmladd(text)
    mymail.addattach([errorFile])
    mymail.send()

def main():

    loadConfig()
    loadAGmapping()
    loadTourmapping()

    files = [f for f in listdir(inFolder) if isfile(join(inFolder, f))]
    for file in files:
        readCSV(inFolder+"/"+file)

    if inputErrorCount > 1:
        handleErrors()

main()
