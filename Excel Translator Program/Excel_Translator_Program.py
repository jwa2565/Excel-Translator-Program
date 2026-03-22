import pandas as pd

df = pd.read_excel("Excel Translator Spreadsheet.xlsx")

import zipfile
import json
from pydantic import BaseModel

import xml.etree.ElementTree as ET

from google import genai

class Translation(BaseModel):
    language: str
    words: list[str]
    
class File(BaseModel):
    fileName: str
    translations_Cells: list[Translation]
    translations_Shapes: list[Translation]
    
def cleanDataframe():
     df["Additional Language 1"] = df["Additional Language 1"].fillna("none")
     df["Additional Language 2"] = df["Additional Language 2"].fillna("none")
     df["Additional Language 3"] = df["Additional Language 3"].fillna("none")
     df["Additional Language 4"] = df["Additional Language 4"].fillna("none")
     
def unzipFiles(fileName):
    zip_file_path = fileName_temp + ".xlsx"
    extract_to_path = fileName_temp + "_English" 
    
    try:
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
          zip_ref.extractall(extract_to_path)
          print(f"Successfully extracted all files from '{zip_file_path}' to '{extract_to_path}'.")
        
    except zipfile.BadZipFile:
        print(f"Error: '{zip_file_path}' is not a valid zip file or is corrupted.")
    except FileNotFoundError:
        print(f"Error: Zip file '{zip_file_path}' not found.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    

def extractCellContents(fileName_temp):

    
    wordList = []

    cellFilePath = fileName_temp + "_English/xl/sharedStrings.xml"

    tree = ET.parse(cellFilePath)
    root = tree.getroot()

    for t in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
        wordList.append(t.text)
    
    
    return wordList

def extractShapeContents(fileName_temp):

    
    wordList = []

    shapeFilePath = fileName_temp + "_English/xl/drawings/drawing1.xml"

    tree = ET.parse(shapeFilePath)
    root = tree.getroot()

    ns = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

    for t in root.findall(".//a:t", ns):   #for t in root.iter("a:t"):
        if t.text:
            wordList.append(t.text)
           # print(t.text)
    
    
    return wordList

def generateWordListToSend(excelFiles):
    wordsToSend = []

    for i in range(len(excelFiles)):  
       # print(excelFiles[i].fileName)
        for j in range(len(excelFiles[i].translations_Cells[0].words)):
            wordsToSend.append(excelFiles[i].translations_Cells[0].words[j])
            #print(excelFiles[i].translations_Cells[0].words[j])

        #print("$$$$$$")
        wordsToSend.append("$$$$$$")
    
        for j in range(len(excelFiles[i].translations_Shapes[0].words)):
            #print(excelFiles[i].translations_Shapes[0].words[j])
            wordsToSend.append(excelFiles[i].translations_Shapes[0].words[j])
    
        
        #print("#########")
        wordsToSend.append("#########")
    
    return wordsToSend

def defaultLanguages():
    defaultLanguages = []
    defaultLanguages.append("Spanish")
    defaultLanguages.append("Arabic")
    defaultLanguages.append("Haitian Creole")
    defaultLanguages.append("Swahili")
    defaultLanguages.append("French")
    defaultLanguages.append("Vietnamese")
    
    return defaultLanguages

def generateLanguageList():
    languageList = defaultLanguages()
    
    for i, rows in df.iterrows():
        if df.iloc[i]["Additional Language 1"] != "none":
            if df.iloc[i]["Additional Language 1"] not in languageList:
                languageList.append(df.iloc[i]["Additional Language 1"])
              
        if df.iloc[i]["Additional Language 2"] != "none":
            if df.iloc[i]["Additional Language 2"] not in languageList:
                languageList.append(df.iloc[i]["Additional Language 2"])
                
        if df.iloc[i]["Additional Language 3"] != "none":
            if df.iloc[i]["Additional Language 3"] not in languageList:
                languageList.append(df.iloc[i]["Additional Language 3"])

        if df.iloc[i]["Additional Language 4"] != "none":
            if df.iloc[i]["Additional Language 4"] not in languageList:
                languageList.append(df.iloc[i]["Additional Language 4"])
    
    return languageList

def generateQuestion(languageList):
    question = "Translate the list of strings into "
    
    for i in range(len(languageList)):
    
        if i != len(languageList) - 1:
            question = question + languageList[i] + ", "
        
        else:
            question = question + "and " + languageList[i] + ". "

    question = question + "Please don't skip any elements in the list. If you can't translate something (such as a person's name), simply return the unmodified string. If you have a string that you can only partially translate, return a single string with both the translated and untranslated parts together. For example, if the string is *I love you, Kirsten* and the language is Spanish, return a single string *Te amo, Kirsten*. It is very important that the resulting word list you give me should be the same size as the word list that I input to you. "
    #question = question + "Please return the strings using the standard roman alphabet. "
    return question

def aiStuff(question, words):
    
    print("Generating Translations...")
    
    client = genai.Client(api_key="NUNYA, WHOOPSSSSSSSS")
    
    response = client.models.generate_content(
        
        model="gemini-2.5-flash-lite",
        contents=[question, words],
        config={
            'response_mime_type': 'application/json',
            'response_schema': list[Translation]
         }

#
        
     )
    
    print("Finished Translating!")
    
    return response.parsed
   
excelFiles: list[File] = []
cleanDataframe()

for i, rows in df.iterrows():
    
    fileName_temp = df.iloc[i]["filename of excel spreadsheet to translate"]
    baseLanguage = "English"
    
    unzipFiles(fileName_temp)
    
    cell = [Translation(language=baseLanguage, words=extractCellContents(fileName_temp))]
    shape = [Translation(language=baseLanguage, words=extractShapeContents(fileName_temp))]

    tempFile = File(fileName=fileName_temp, translations_Cells=cell, translations_Shapes=shape)
    excelFiles.append(tempFile)


wordList = generateWordListToSend(excelFiles)
languageList = generateLanguageList()
question = generateQuestion(languageList)

    
for i in range(len(wordList)): #ai gets confused sometimes and merges words together
    wordList[i] = (wordList[i]).strip()
    wordList[i] = " " + wordList[i] + " "
    


my_translations: list[Translation] = aiStuff(question, wordList)

#file = open("results.txt", "w")
    
#for i in range(len(my_translations)):
#    file.write(str(my_translations[i]))
#file.close()

with open("results.txt", "w", encoding="utf-8") as file:
    for i in range(len(my_translations)):
        #file.write(str(my_translations[i]))
        file.write(my_translations[i].language)
        file.write("\n")
        for j in range(len(my_translations[i].words)):
            toWrite = str(j) + ": " + my_translations[i].words[j] + "\n"
            file.write(toWrite)
       
#counter = 0
#for i in range(len(excelFiles)):

#for i in range(len(my_translations)):

fileName = "translation_" + "English" + ".txt"
file = open(fileName, "w")

for i in range(len(wordList)):
    
    toWrite = str(i) + ": " + wordList[i] + "\n"
    file.write(toWrite)
        
file.close()
##################
for i in range(len(my_translations)):
    
    fileName = "translation_" + my_translations[i].language + ".txt"
    with open(fileName, "w", encoding="utf-8") as file:
        for j in range(len(my_translations[i].words)):
            toWrite = str(j) + ": " + my_translations[i].words[j] + "\n"
            file.write(toWrite)
    
    
    
    
    
    

    
    
    





    

