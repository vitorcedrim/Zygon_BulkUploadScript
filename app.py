import zipfile, glob, os, openpyxl, re
from PySimpleGUI.PySimpleGUI import WIN_CLOSED, popup
from re import search
from openpyxl import Workbook
import PySimpleGUI as sg

varTaxonomy = ""
varCreativeType = ""
varClickTroughURL = ""
varTrackingURL1 = ""
varTrackingURLType1 = ""
varTrackingURL2 = ""
varTrackingURLType2 = ""

def displayCreativeWorkflow():
    wb = Workbook()
    ws = wb.active

    arrayFileList = [] #Armazena os arquivos de criativo que estão dentro do ZIP.

    varCreativeQuantity = 0 #Quantidade de criativos dentro do ZIP, utilizado como condição para interromper o loop.
    varCreativeNumber = 0 #Número/índice do criativo sendo tratado no loop atual.
    varFileName = "" #Nome do arquivo de criativo sendo tratado no loop atual.
    varFileSize = "" #Tamanho em pixels do criativo sendo tratado no loop atual.
    varCreativeStatus = "Active"
    varCreativeAdChoice = "Yes"
    varCreativeName = "" #Nome do arquivo de criativo sendo tratado no loop atual.



    file = glob.glob("*.zip") #Busca por arquivos ZIP na pasta raiz.
    with zipfile.ZipFile(file[0], "r") as zip_file:
        arrayFileList = zip_file.namelist() #Armazena os arquivos de criativo que estão dentro do ZIP dentro de arrayFileList.

    varCreativeQuantity = len(arrayFileList)

    def createDisplayTemplate():
        ws.title = "Display - Hosted"

        ws['A1'] = "Creative Name*"
        ws['B1'] = "Image File Name*"
        ws['C1'] = "Image File Secure URL"
        ws['D1'] = "Image File Un-secure URL"
        ws['E1'] = "Ad Size*"
        ws['F1'] = "Status*"
        ws['G1'] = "Enable Ad Choices*"
        ws['H1'] = "Creative Category"
        ws['I1'] = 'Custom Id ("," Separated)'
        ws['J1'] = "Click Through URL*"
        ws['K1'] = "Landing Page URL"
        ws['L1'] = "Tracking URL 1"
        ws['M1'] = "Tracking URL TYPE 1"
        ws['N1'] = "Tracking URL 2"
        ws['O1'] = "Tracking URL TYPE 2"
        ws['P1'] = "Tracking URL 3"
        ws['Q1'] = "Tracking URL TYPE 3"
        ws['R1'] = "Tracking URL 4"
        ws['S1'] = "Tracking URL TYPE 4"
        ws['T1'] = "Tracking URL 5"
        ws['U1'] = "Tracking URL TYPE 5"
        ws['V1'] = "Creative Id"
        ws['W1'] = "Line Id"
        ws['X1'] = "Line Name"

    def writeDataDisplay():
        ws[f'A{varCreativeNumber + 2}'] = varCreativeName
        ws[f'B{varCreativeNumber + 2}'] = varFileName
        ws[f'E{varCreativeNumber + 2}'] = varFileSize
        ws[f'F{varCreativeNumber + 2}'] = varCreativeStatus
        ws[f'G{varCreativeNumber + 2}'] = varCreativeAdChoice
        ws[f'J{varCreativeNumber + 2}'] = varClickTroughURL
        ws[f'L{varCreativeNumber + 2}'] = varTrackingURL1
        ws[f'M{varCreativeNumber + 2}'] = varTrackingURLType1
        ws[f'N{varCreativeNumber + 2}'] = varTrackingURL2
        ws[f'O{varCreativeNumber + 2}'] = varTrackingURLType2
    

    createDisplayTemplate()


    while varCreativeNumber < varCreativeQuantity:
        varFileName = arrayFileList[varCreativeNumber] #Coloca o nome do arquivo baseado no arquivo que está sendo tratado no loop atual.
        varCreativeSize = varCreativeNumber

        if re.search(r"1080x1080", varFileName):
            varFileSize = "1080x1080"
        elif re.search(r"160x600", varFileName):
            varFileSize = "160x600"
        elif re.search(r"600x600", varFileName):
            varFileSize = "600x600"
        elif re.search(r"300x600", varFileName):
            varFileSize = "300x600"
        elif re.search(r"970x250", varFileName):
            varFileSize = "970x250"
        elif re.search(r"970x500", varFileName):
            varFileSize = "970x500"
        elif re.search(r"120x60", varFileName):
            varFileSize = "120x60"
        elif re.search(r"300x250", varFileName):
            varFileSize = "300x250"
        elif re.search(r"700x500", varFileName):
            varFileSize = "700x500"
        elif re.search(r"640x1136", varFileName):
            varFileSize = "640x1136"
        elif re.search(r"320x50", varFileName):
            varFileSize = "320x50"
        elif re.search(r"320x568", varFileName):
            varFileSize = "320x568"
        elif re.search(r"728x90", varFileName):
            varFileSize = "728x90"
        elif re.search(r"728x500", varFileName):
            varFileSize = "728x500"
        elif re.search(r"1030x60", varFileName):
            varFileSize = "1030x60"
        elif re.search(r"535x45", varFileName):
            varFileSize = "535x45"
        elif re.search(r"580x60", varFileName):
            varFileSize = "580x60"
        elif re.search(r"255x60", varFileName):
            varFileSize = "255x60"
        elif re.search(r"1440x1024", varFileName):
            varFileSize = "1440x1024"
        elif re.search(r"600x450", varFileName):
            varFileSize = "600x450"
        elif re.search(r"1440x810", varFileName):
            varFileSize = "1440x810"
        elif re.search(r"360x405", varFileName):
            varFileSize = "360x405"
        elif re.search(r"1920x1080", varFileName):
            varFileSize = "1920x1080"
        elif re.search(r"1080x1440", varFileName):
            varFileSize = "1080x1440"
        elif re.search(r"100x160", varFileName):
            varFileSize = "100x160"
        elif re.search(r"1080x1920", varFileName):
            varFileSize = "1080x1920"
        elif re.search(r"416x216", varFileName):
            varFileSize = "416x216"
        elif re.search(r"468x263", varFileName):
            varFileSize = "468x263"
        elif re.search(r"970x90", varFileName):
            varFileSize = "970x90"
        elif re.search(r"320x100", varFileName):
            varFileSize = "320x100"
        elif re.search(r"320x480", varFileName):
            varFileSize = "320x480"
        
        varCreativeName = varTaxonomy + varFileSize
        print(varCreativeName)
        writeDataDisplay() #Escreve os dados na planilha.
        varCreativeNumber += 1 #Avança o índice de criativo a ser tratado.
        
    wb.save('bulk.xlsx') #Salva o XLSX final.

def videoCreativeWorkflow():
    wb = Workbook()
    ws = wb.active

    arrayFileList = [] #Armazena os arquivos de criativo que estão dentro do ZIP.

    varCreativeQuantity = 0 #Quantidade de criativos dentro do ZIP, utilizado como condição para interromper o loop.
    varCreativeNumber = 0 #Número/índice do criativo sendo tratado no loop atual.
    varFileName = "" #Nome do arquivo de criativo sendo tratado no loop atual.
    varFileSize = "" #Tamanho em pixels do criativo sendo tratado no loop atual.
    varCreativeStatus = "Active"
    varCreativeAdChoice = "Yes"
    varCreativeName = "" #Nome do arquivo de criativo sendo tratado no loop atual.



    file = glob.glob("*.zip") #Busca por arquivos ZIP na pasta raiz.
    with zipfile.ZipFile(file[0], "r") as zip_file:
        arrayFileList = zip_file.namelist() #Armazena os arquivos de criativo que estão dentro do ZIP dentro de arrayFileList.

    varCreativeQuantity = len(arrayFileList)

    def createVideoTemplate():
        ws.title = "Video - Hosted"

        ws['A1'] = "Creative Name*"
        ws['B1'] = "Video File Name*"
        ws['C1'] = "Video File Secure URL"
        ws['D1'] = "Video File Un-secure URL"
        ws['E1'] = "Status*"
        ws['F1'] = "Enable Ad Choices*"
        ws['G1'] = "Creative Category"
        ws['H1'] = 'Custom Id ("," Separated)'
        ws['I1'] = 'Click Through URL*'
        ws['J1'] = "Landing Page URL"
        ws['K1'] = "Clearcast Clock Number"
        ws['L1'] = "Tracking URL 1"
        ws['M1'] = "Tracking URL TYPE 1"
        ws['N1'] = "Tracking URL 2"
        ws['O1'] = "Tracking URL TYPE 2"
        ws['P1'] = "Tracking URL 3"
        ws['Q1'] = "Tracking URL TYPE 3"
        ws['R1'] = "Tracking URL 4"
        ws['S1'] = "Tracking URL TYPE 4"
        ws['T1'] = "Tracking URL 5"
        ws['U1'] = "Tracking URL TYPE 5"
        ws['V1'] = "Creative Id"
        ws['W1'] = "Line Id"
        ws['X1'] = "Line Name"

    def writeDataVideo():
        ws[f'A{varCreativeNumber + 2}'] = varCreativeName
        ws[f'B{varCreativeNumber + 2}'] = varFileName
        ws[f'E{varCreativeNumber + 2}'] = varCreativeStatus
        ws[f'F{varCreativeNumber + 2}'] = varCreativeAdChoice
        ws[f'I{varCreativeNumber + 2}'] = varClickTroughURL
        ws[f'L{varCreativeNumber + 2}'] = varTrackingURL1
        ws[f'M{varCreativeNumber + 2}'] = varTrackingURLType1
        ws[f'N{varCreativeNumber + 2}'] = varTrackingURL2
        ws[f'O{varCreativeNumber + 2}'] = varTrackingURLType2
    

    createVideoTemplate()


    while varCreativeNumber < varCreativeQuantity:
        varFileName = arrayFileList[varCreativeNumber] #Coloca o nome do arquivo baseado no arquivo que está sendo tratado no loop atual.
        varCreativeName = varTaxonomy + varFileSize
        writeDataVideo() #Escreve os dados na planilha.
        varCreativeNumber += 1 #Avança o índice de criativo a ser tratado.

    wb.save('bulk.xlsx') #Salva o XLSX final.

def taxonomyWindow():
    sg.theme("Reddit")
    layout = [
        [sg.Text("Insira a taxonomia Zygon para criativo (Sem dimensões).")],
        [sg.Input(key="taxonomy")],
        [sg.Button("Continuar")]
    ]
    return sg.Window("Taxonomia", layout=layout,finalize=True)

def creativeTypeWindow():
    sg.theme("Reddit")
    layout = [
        [sg.Text("Escolha o tipo de criativo que deseja fazer upload.")],
        [sg.Radio('Display', "radioCreativeType", default=True, key="DisplayAd")], [sg.Radio('Vídeo', "radioCreativeType", default=False, key="VideoAd")],
        [sg.Button("Continuar")]
    ]
    return sg.Window("Tipo de Criativo", layout=layout, finalize=True)

def extraInfosWindow():
    sg.theme("Reddit")
    layout = [
        [sg.Text("Insira as informações extras.")],
        [sg.Text("Click Through URL (Lembre de parametrizar)")],[sg.Input(key="clickTroughURL")],
        [sg.Text("Tracking URL 1 (Opcional)")],[sg.Input(key="trackingURL1")],
        [sg.Text("Tracking URL Type 1 (Opcional)")],[sg.Input(key="trackingURLType1")],
        [sg.Text("Tracking URL 2 (Opcional)")],[sg.Input(key="trackingURL2")],
        [sg.Text("Tracking URL Type 2 (Opcional)")],[sg.Input(key="trackingURLType2")],
        [sg.Button("Criar planilha de bulk upload!")]
    ]
    return sg.Window("Extras", layout=layout,finalize=True)

def finalWindow():
    sg.theme("Reddit")
    layout = [
        [sg.Text("Script criado com muito carinho para a equipe de operações Zygon por:")],
        [sg.Text("Vitor Cedrim")],
        [sg.Button("Finalizar")]
    ]
    return sg.Window("Obrigado!", layout=layout,finalize=True, element_justification='center')

varTaxonomyWindow = taxonomyWindow()
varCreativeTypeWindow = None
varExtraInfosWindow = None
varFinalWindow = None
while True:
    window,event,values = sg.read_all_windows()

    if window == varTaxonomyWindow and event == sg.WIN_CLOSED:
        break

    if window == varTaxonomyWindow and event == "Continuar" and values["taxonomy"] != "":
        varTaxonomy = values["taxonomy"]
        varCreativeTypeWindow = creativeTypeWindow()
        varTaxonomyWindow.hide()
    if window == varCreativeTypeWindow and event == sg.WIN_CLOSED:
        break
    if window == varCreativeTypeWindow and event == "Continuar":
        if values["DisplayAd"] == True:
            varCreativeType = "Display"
            varCreativeTypeWindow.hide()
            varExtraInfosWindow = extraInfosWindow()
        elif values["VideoAd"] == True:
            varCreativeType = "Video"
            varCreativeTypeWindow.hide()
            varExtraInfosWindow = extraInfosWindow()
    if window == varExtraInfosWindow and event == sg.WIN_CLOSED:
        break
    if window == varExtraInfosWindow and event == "Criar planilha de bulk upload!":
        if varCreativeType == "Display":
            varClickTroughURL = values["clickTroughURL"]
            varTrackingURL1 = values["trackingURL1"]
            varTrackingURLType1 = values["trackingURLType1"]
            varTrackingURL2 = values["trackingURL2"]
            varTrackingURLType2 = values["trackingURLType2"]
            displayCreativeWorkflow()
            varExtraInfosWindow.hide()
            varFinalWindow = finalWindow()
        elif varCreativeType == "Video":
            varClickTroughURL = values["clickTroughURL"]
            varTrackingURL1 = values["trackingURL1"]
            varTrackingURLType1 = values["trackingURLType1"]
            varTrackingURL2 = values["trackingURL2"]
            varTrackingURLType2 = values["trackingURLType2"]
            videoCreativeWorkflow()
            varExtraInfosWindow.hide()
            varFinalWindow = finalWindow()
    if window == varFinalWindow and event == "Finalizar":
        break
    elif window == varFinalWindow and event == sg.WIN_CLOSED:
        break