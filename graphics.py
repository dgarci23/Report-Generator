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

    graphicsWorksheet = excelData.active

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

def getCriteriaPerMonth(criteriaName, criteriaCols, data, indexes):

    graphicsInformation[criteriaName] = {}

    for sheet in indexes.keys():

        currentSheetData = data[sheet]

        graphicsInformation[criteriaName][sheet] = [currentSheetData[col + str(indexes[sheet])].value for col in criteriaCols ]

    return graphicsInformation

def editCriteriaAllMonths(criteriaName):

    criteriaDict = graphicsInformation[criteriaName]

    dataPoints = len(criteriaDict[list(criteriaDict.keys())[0]])

    editList = [0 for i in range(dataPoints)]

    for key in criteriaDict.keys():

        index = 0
        for element in criteriaDict[key]:

            editList[index] += element

    graphicsInformation[criteriaName] = editList

def addCriteriaPerMonth(criteriaName):



# -------- Initializing information

worksheets = getWorksheets()

parametersIndexes = getParametersIndexes()

excelData = load_workbook(fileName, data_only=True, read_only=True)

graphicsSheet, graphicsSpreadsheet = createGraphicsWorksheet()

indexes = getIndexes(excelData, worksheets)

# --------- Processing the data and creating the tables

graphicsInformation = {}

criteriaColsPerMonth = {
    "Impactos": ["K"],
    "Tier": ["O", "P", "Q"],
    "Valor Publicitario": ["L"],
    "Valor Informativo": ["M"],
    "Favorabilidad Mediatica": ["X", "Y", "Z"],
    "Quote de Vocero" : ["V"],
    "Presencia de Vocero": ["T"],
    "Mencion de Marca": ["R"]
}

for criteria in criteriaColsPerMonth.keys():

    getCriteriaPerMonth(criteria, criteriaColsPerMonth[criteria], excelData, indexes)

editCriteriaAllMonths("Favorabilidad Mediatica")

print(graphicsInformation)


# --------- Saving files

graphicsSpreadsheet.save("Finished.xlsx")




