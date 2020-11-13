from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject
from pathlib import Path
import datetime
import csv

PROPERTY_CONTROL_TEMPLATE = 'propertyControlTemplate.pdf'
PROPERTY_CONTROL_OUTPUT = 'propertyControlForm.pdf'

ASSET_CSV = 'assets.csv'

pdf_path = (
    Path.home()
    / "Desktop"
    / "PropControlFill"
    / PROPERTY_CONTROL_TEMPLATE
)

print("!!YOU MUST UPDATE ITAM FOR ALL ASSETS BEFORE CONTINUING!!")

def set_need_appearances_writer(writer: PdfFileWriter):
    try:
        catalog = writer._root_object

        if "/AcroForm" not in catalog:
            writer._root_object.update({
                NameObject("/AcroForm"): IndirectObject(len(writer._objects), 0, writer)})

        need_appearances = NameObject("/NeedAppearances")
        writer._root_object["/AcroForm"][need_appearances] = BooleanObject(True)
        return writer

    except Exception as e:
        print('set_need_appearances_writer() catch : ', repr(e))
        return writer

pdf = PdfFileReader(open(PROPERTY_CONTROL_TEMPLATE, "rb"), strict=False)
if "/AcroForm" in pdf.trailer["/Root"]:
    pdf.trailer["/Root"]["/AcroForm"].update(
        {NameObject("/NeedAppearances"): BooleanObject(True)})

def getAssetInfo (file):
    assetInfo = input("Please enter the Asset Tag Number: ")
    with open(file, 'r') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if str(assetInfo) in row[0]:
                return row

def getLocation(location):
    location = location.split()
    if 'West' in location:
        return location[0]
    elif 'East' in location:
        return location[0]
    elif 'Downtown' in location:
        return location[0]

def buildDictionary (assetInfo):
    date = datetime.datetime.now()
    dateNow = str(date.month) + "-" + str(date.day) + "-" + str(date.year)

    dictionary = {}

    dictionary['Date'] = dateNow
    dictionary['PCC Number'] = assetInfo[0]
    dictionary['Serial Number'] = assetInfo[3]
    dictionary['Item Description'] = assetInfo[1] + " " + assetInfo[2]
    location = getLocation(assetInfo[4])
    dictionary['CampusSite'] = location
    dictionary['BuildingRoom'] = assetInfo[5]
    dictionary['Date Available for Pickup'] = dateNow
    dictionary['CampusSite_2'] = input("What is the gaining campus?: ")
    dictionary['BuildingRoom_2'] = input("What is the gaining Room?: ")
    dictionary['Gaining Custodian'] = input("Who is the Gaining Custodian?: ")
    return dictionary    

pdfOut = PdfFileWriter()
set_need_appearances_writer(pdfOut)
if "/AcroForm" in pdfOut._root_object:
    pdfOut._root_object["/AcroForm"].update(
        {NameObject("/NeedAppearances"): BooleanObject(True)})

itamFile = ASSET_CSV
assetInfo = getAssetInfo(itamFile)

dictionary = buildDictionary(assetInfo)

pdfOut.addPage(pdf.getPage(0))
pdfOut.updatePageFormFieldValues(pdfOut.getPage(0), dictionary)

outputStream = open(PROPERTY_CONTROL_OUTPUT, "wb")
pdfOut.write(outputStream)

print("!!PLEASE REVIEW FORM FOR DISCREPENCIES BEFORE PRINTING!!")
