# IMPORT ALL LIBRARIES
# Import os library - to get list of file from folder
import os
# Import re library - to use regular expression to receive wanted strings
import re
# Install the pdfminer.six package and use it - to read value from pdf file
from pdfminer.high_level import extract_text
# datetime is by default installed - no need to install it from console.
from datetime import datetime
# Install the xlswriter package and use it - to create the output Excel file
import xlsxwriter

# ENVIRONMENT VARIABLES
# Path to folder with input pdf. By default - it is folder named 'INPUT' in a same location as this script.
inputFolder = 'INPUT'
# Path to folder where output excel will be stored.
outputFolder = 'OUTPUT'
# Get the timestamp of now to prepare unique output Excel file name
now = datetime.now()
# Make string out of it and replace the sign does not allow for Excel file name
now = str(now).replace(':', '.')

# Start Excel file
outputWorkbook = xlsxwriter.Workbook(outputFolder + '/output' + now + '.xlsx')
# Create new worksheet to which data will be provided
outputWorksheet = outputWorkbook.add_worksheet(name='OUTPUT')

# Declare the var with which the first empty row in output excel will be track
rowNum = 1
# Do in loop for each element in 'inputFolder' location
for filename in os.listdir(inputFolder):
    # Do only if element is a file and a file ends with .pdf extension
    pdfFile = os.path.join(inputFolder, filename)
    if os.path.isfile(pdfFile) & pdfFile.endswith('.pdf'):
        try:
            # Extract all string data from page one of PDF
            PDFtext = extract_text(pdfFile, maxpages=1)
        except:
            # When not able to receive data from pdf for any reason - continue the loop and try with next pdf from location
            continue
        # Regex the invoice number
        invoiceNumber_regexPattern = re.compile(r"INVOICE #(.*) ")
        invoiceNumber = invoiceNumber_regexPattern.findall(PDFtext)
        # Regex the invoice date
        invoiceDate_regexPattern = re.compile(r"DATE: (.*) ")
        invoiceDate = invoiceDate_regexPattern.findall(PDFtext)
        # Provide headers.
        # POSSIBLE IMPROVEMENT - check if the header were provided earlier.
        outputWorksheet.write(0, 0, 'invoice #')
        outputWorksheet.write(0, 1, 'Date')
        outputWorksheet.write(0, 2, 'File name')
        # Provide read values
        try:
            outputWorksheet.write(rowNum, 0, invoiceNumber[0])
        except IndexError:
            # If IndexError occur, it means that the previous regex value did not founded value with given pattern.
            outputWorksheet.write(rowNum, 0, 'ERROR: No invoice # detected')
        try:
            outputWorksheet.write(rowNum, 1, invoiceDate[0])
        except IndexError:
            outputWorksheet.write(rowNum, 1, 'ERROR: No date detected')
        outputWorksheet.write(rowNum, 2, filename)
        rowNum += 1

# Close the Excel file
try:
    outputWorkbook.close()
except:
    # If error occur - create the output folder and try to save excel once again
    os.makedirs(outputFolder)
    outputWorkbook.close()
