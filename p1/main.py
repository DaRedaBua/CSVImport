#v1.0.2

import xml.etree.ElementTree as ET
from os import listdir, rename
from os.path import isfile, join
import datetime
import Bimail
import calendar
import log
from log import log, loadLogs, getStackTrace, getAllTrace
from pathlib import Path
import codecs

ag_mapping = dict()
tour_mapping = dict()

againLines = []
msgs = []
inputErrorCount = 0

todaySTR = datetime.datetime.now().strftime('%Y-%m-%d_%f')[:-3]
auftrag = 0
recipients = []

outFolder = ""
inFolder = ""
errFolder = ""
backupFolder = ""
AgMappingPath = ""
TourMappingPath = ""
sourcePath = ""

def readCSV(path):
    log(400, "readCSV(path) - path:", path)

    global inputErrorCount
    global auftrag

    with open(path, 'r') as f:
        lines = f.readlines()

        for line in lines:
            log(405, "line: ", line)
            prnt = True

            tree = ET.parse(sourcePath)
            root = tree.getroot()

            NK = root.find('NK')
            AK = NK.find('AK')
            SD = AK.find('SD')
            TOUR = AK.find('TOURVORGABE')
            RSTK = SD.find('RSTK')

            lineElems = line.split(";")
            log(410, "lineElems: ", lineElems)
        #Datums
            dat = lineElems[1]
            if len(lineElems[0]) < 2:
                dat = dat + "0"
            dat = dat + lineElems[0]

            #erster, letzter
            fom = dat + "01"
            lom = dat + str(calendar.monthrange(int(lineElems[1]), int(lineElems[0]))[1])

            log(415, "FirstOfMonth: ", fom)
            log(416, "LasstOfMonth: ", lom)

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
                log(425, "AG " + lineElems[6] + " nicht in AGmapping gefunden, prnt: ", False)
                inputErrorCount = inputErrorCount + 1
                againLines.append(line)
                text = "Zeile: " + str(inputErrorCount) + " - Konnte Auftraggeber " + lineElems[6] + " nicht finden"
                msgs.append(text)
                prnt = False
            else:
                log(426, "AG-Mapping <ADR-AG>,<ADR-ABS> : key,value", lineElems[6] + " : " + ag_mapping[lineElems[6]])
                adrag = AK.find('ADR-AG')
                agnr = adrag.find('NR')
                agnr.text = ag_mapping[lineElems[6]]

                adrabs = SD.find('ADR-ABS')
                absnr = adrabs.find('NR')
                absnr.text = ag_mapping[lineElems[6]]


        #EMP/Entladestelle + INFO
            #Prnt wird zweimal abgefragt WEIL wenn prnt bereits FALSE wird sowieso nicht gedruckt.
            #Bei UNGÜLTIG wird verhindert, dass die Zeile ein zweites mal hinzugefügt wird,
            #bei GÜLTIG wird verhindert, dass obwohl ERROR danach gesucht wird
            #UNGÜLTIG
            if lineElems[5] not in tour_mapping and prnt:
                log(435, "Tour " + lineElems[5] + " nicht in Tourmapping gefunden, prnt: ", False)
                inputErrorCount = inputErrorCount + 1
                againLines.append(line)
                text = "Zeile: " + str(inputErrorCount) + " - Konnte Tour " + lineElems[5] + " nicht finden"
                msgs.append(text)
                prnt = False
            elif prnt:
                #GÜLTIG
                adremp = SD.find('ADR-EMP')
                #Z'FUAS
                if tour_mapping[lineElems[5]][1] == '9999':
                    log(440, "Tour Z'Fuas eingeben", "")
                    name = ET.Element('NAME1')
                    name.text = tour_mapping[lineElems[5]][2]
                    adremp.append(name)
                    log(441, "<NAME1>: ", name)
                    ort = ET.Element('ORT1')
                    ort.text = tour_mapping[lineElems[5]][3]
                    adremp.append(ort)
                    log(442, "<ORT1>: ", ort)
                    land = ET.Element('LAND')
                    land.text = tour_mapping[lineElems[5]][4]
                    adremp.append(land)
                    log(443, "<LAND>: ", land)
                    isoland = ET.Element('ISOLAND')
                    isoland.text = tour_mapping[lineElems[5]][5]
                    adremp.append(isoland)
                    log(444, "<ISOLAND>: ", isoland)
                    plz = ET.Element('PLZ')
                    plz.text = tour_mapping[lineElems[5]][6]
                    adremp.append(plz)
                    log(445, "<PLZ>: ", plz)
                    strasse = ET.Element('STRASSE')
                    strasse.text = tour_mapping[lineElems[5]][7]
                    adremp.append(strasse)
                    log(446, "<STRASSE>: ", strasse)
                    hausnr = ET.Element('HAUSNR')
                    hausnr.text = tour_mapping[lineElems[5]][8]
                    adremp.append(hausnr)
                    log(447, "<HAUSNR>: ", hausnr)
                #INFO
                    info = SD.find('INFO')
                    print(info)
                    info1 = info.find('INFO1')
                    print(info1)

                    info1.text = tour_mapping[lineElems[5]][2]
                    log(448, "<INFO1>: ", lineElems[5])
                #HINTERLEGT
                else:
                    log(450, "Adresse in Carlo Mapping gefunden", "")
                    nr = ET.Element('NR')
                    nr.text = tour_mapping[lineElems[5]][1]
                    adremp.append(nr)
                    log(451, "CarLo-GP-NR:", nr)

        #GEWICHT
            PO = SD.find('PO')
            POFD = PO.find('POFD')
            tgewicht = POFD.find('TGEWICHT')
            tgewicht.text = lineElems[7]
            ugewicht = POFD.find('UGEWICHT')
            ugewicht.text = lineElems[7]
            fgewicht = POFD.find('FGEWICHT')
            fgewicht.text = lineElems[7]
            log(460, "<T/F/U-Gewicht>: ", lineElems[7])

        #EXTNR
            extnr = AK.find('EXTNR')
            extnr.text = "Tour " + lineElems[5]
            log(470, "<EXTNR>: ", lineElems[5])

        #OUTPUT
            if prnt:
                log(495, "Schreibe in Baum...", "")
                auftrag = auftrag + 1
                file = outFolder + "/AT_" + todaySTR + "_" + str(auftrag) + ".xml"
                tree.write(file)
                log(496, "ERFOLG, Datei: ", file)


def loadConfig():
    log(100, "loadConfig() - ", "opening ./config.csv")

    global recipients
    global inFolder
    global outFolder
    global errFolder
    global backupFolder
    global AgMappingPath
    global TourMappingPath
    global sourcePath

    with open('config.csv', 'r') as f:
        lines = f.readlines()
        print(lines)
        for x in range(len(lines)):
            log(105, "Config-Line: ", x)

            lineElems = lines[x].split(';')
            log(106, "Line-Elems: ", lineElems)

            if x == 0:
                for y in range(1, len(lineElems)):
                    log(110, "recipients.append(lineElem)", lineElems[y])
                    recipients.append(lineElems[y])

            if x == 1:
                if lineElems[2] == '1':
                    log(115, "Creating inFolder", "")
                    Path(lineElems[1]).mkdir(parents=True, exist_ok=True)
                log(116, "inFolder: ", lineElems[1])
                inFolder = lineElems[1]


            if x == 2:
                if lineElems[2] == '1':
                    log(120, "Creating outFolder", "")
                    Path(lineElems[1]).mkdir(parents=True, exist_ok=True)
                log(121, "outFolder: ", lineElems[1])
                outFolder = lineElems[1]

            if x == 3:
                if lineElems[2] == '1':
                    log(125, "Creating errFolder", "")
                    Path(lineElems[1]).mkdir(parents=True, exist_ok=True)
                log(126, "errFolder: ", lineElems[1])
                errFolder = lineElems[1]

            if x == 4:
                if lineElems[2] == '1':
                    log(125, "Creating backupFolder", "")
                    Path(lineElems[1]).mkdir(parents=True, exist_ok=True)
                log(126, "baclupFolderFolder: ", lineElems[1])
                backupFolder = lineElems[1]

            if x == 5:
                log(130, "AgMappingPath: ", lineElems[1])
                AgMappingPath = lineElems[1]

            if x == 6:
                log(135, "TourMappingPath: ", lineElems[1])
                TourMappingPath = lineElems[1]

            if x == 7:
                log(140, "sourcePath: ", lineElems[1])
                sourcePath = lineElems[1]


def loadAGmapping():
    log(200, "loadAGmapping()", "")
    with open(AgMappingPath, 'r') as f:
        lines = f.readlines()

        for line in lines:
            log(205, "line: ", line)
            l = line.split(';')
            ag_mapping[l[0]] = l[1]
            log(210, "ag_mapping Key/Value: ", l[0] + " - " + l[1])

    log(290, "ag_mapping: ", ag_mapping)


def loadTourmapping():
    log(300, "loadTourmapping()", "")
    with open(TourMappingPath, 'r') as f:
        lines = f.readlines()

        for line in lines:
            log(305, "line: ", line)
            l = line.split(';')
            # Ganze zeile als value, erster eintrag als key
            tour_mapping[l[0]] = l
            log(310, "tour_mapping Key/Value: ", l[0] + " - " + l[1])

    log(390, "tour_mapping: ", tour_mapping)


def handleErrors():
    log(500, "handleErrors()", "")

#Schreibe Neue CSV
    errorFile = errFolder+"/err"+todaySTR+".csv"
#Schicke StackTrace und AllTrace mit
    stackFile = errFolder+"/stacktrace_"+todaySTR+".txt"
    allFile = errFolder+"/alltrace_"+todaySTR+".txt"

    log(501, "errorFile: ", errorFile)
    log(502, "stackFile: ", stackFile)
    log(503, "allFile: ", allFile)

    log(599, "log hört hier auf - SendeMail...", "")

    with open(errorFile, 'w') as w:
        w.writelines(againLines)
        w.close()

    with open(stackFile, 'w') as w:
        w.write(getStackTrace())
        w.close()

    with open(allFile, 'w') as w:
        w.write(getAllTrace())
        w.close()

    subj = "Jäger Importe"
    if inputErrorCount > 0:
        text = "Einige Zeilen konnten nicht importiert werden.<br>Diese Zeilen wurden in der angehängten Datei herausgefiltert - der Rest ist importiert worden.<br><br>Zeile <b>(der neuen Datei)</b> + Begründung: <br><br>"

    else:
        text = "Umwandlung in XML erfolgreich!"

    for msg in msgs:
        text = text + msg + "<br>\n"

    mymail = Bimail.Bimail(subj, recipients)
    mymail.htmladd(text)
    mymail.addattach([errorFile, stackFile, allFile])
    mymail.send()


def main():
    loadLogs("logs.csv")

    log("000", "Programm startet", "")

    loadConfig()
    loadAGmapping()
    loadTourmapping()

    files = [f for f in listdir(inFolder) if isfile(join(inFolder, f))]
    for file in files:
        log("010", "Öffne CSV: ", file)
        readCSV(inFolder+"/"+file)
        rename(inFolder+"/"+file, backupFolder+"/"+file)

    handleErrors()

main()
