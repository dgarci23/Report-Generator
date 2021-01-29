from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from selenium import webdriver
from PIL import Image
from docx import Document
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from shutil import copyfile

import glob
import os
import time
import subprocess
import sys
import datetime
import imaplib
import email
import smtplib
import ssl

fileName = "050121_E24_Reporte de Gestion2020_Oficial.xlsx"

def createGraphicsWorksheet():

    excelData = Workbook()

    graphicsWorksheet = excelData.create_sheet("Graficas")

    return graphicsWorksheet, excelData

def loadWorkbook(fileName):

    return load_workbook(fileName, data_only=True, read_only=True)
    
def getWorksheets():

    worksheets = []

    while "Exit" not in worksheets:

        print("Insert the worksheet name or type Exit to continue: ")

        worksheet = input()

        worksheets.append(worksheet)

    del worksheets[-1]

    return worksheets

def getParametersIndexes():

    parametersIndexes = {
        "Impactos": [1,1],
        "Tier": [4,1],
        "Valor Publicitario": [9,1],
        "Valor Informativo": [12,1],
        "Favorabilidad Mediatica": [15,1],
        "Quote de Vocero": [18,1],
        "Presencia de Vocero": [21,1],
        "Mencion de Marca": [24,1],
        "Tipo de Medio": [27,1],
        "Titular de Gestion": [30,1],
        "Tipo de Gestion": [33,1]
    }

    return parametersIndexes

def getIndexes(data, worksheets):

    indexes = {}

    for month in worksheets:

        index = 9

        while (True):

            if (data[month].cell(index,11).value == None):

                indexes[month] = index - 1

                break

            index += 1

    return indexes
    
def getImpactos(data, indexes):

    impactos = {}

    for sheet in indexes.keys():

        currentSheetData = data[sheet]  

        impactos[sheet] = currentSheetData["K" + str(indexes[sheet])].value 

    return impactos 

def addImpactos(impactos):

    row = parametersIndexes["Impactos"][1]
    col = parametersIndexes["Impactos"][0]

    graphicsSheet.cell(row,col).value = "Mes"
    graphicsSheet.cell(row,col+1).value = "Impactos"

    row += 1

    for month in worksheets:
        graphicsSheet.cell(row,col).value = month
        graphicsSheet.cell(row,col+1).value = impactos[month]
        row += 1

def getTier(data, indexes):

    tier = {}

    for sheet in indexes.keys():

        currentSheetData = data[sheet]

        tier[sheet] = [currentSheetData["O" + str(indexes[sheet])].value, currentSheetData["P" + str(indexes[sheet])].value, currentSheetData["Q" + str(indexes[sheet])].value]

    return tier

def addTier(tier):

    row = parametersIndexes["Tier"][1]
    col = parametersIndexes["Tier"][0]

    graphicsSheet.cell(row,col).value = "Mes"
    graphicsSheet.cell(row,col+1).value = "Tier 1"
    graphicsSheet.cell(row,col+2).value = "Tier 2"
    graphicsSheet.cell(row,col+3).value = "Tier 3"

    row += 1

    for month in worksheets:
        graphicsSheet.cell(row,col).value = month
        graphicsSheet.cell(row,col+1).value = tier[month][0]
        graphicsSheet.cell(row,col+2).value = tier[month][1]
        graphicsSheet.cell(row,col+3).value = tier[month][2]
        row += 1



# -------- Initializing information

worksheets = getWorksheets()

parametersIndexes = getParametersIndexes()

excelData = load_workbook(fileName, data_only=True, read_only=True)

graphicsSheet, graphicsSpreadsheet = createGraphicsWorksheet()

indexes = getIndexes(excelData, worksheets)

# --------- Processing the data and creating the tables

impactos = getImpactos(excelData, indexes)
addImpactos(impactos)

tier = getTier(excelData, indexes)
addTier(tier)


# --------- Saving files

graphicsSpreadsheet.save("Finished.xlsx")




