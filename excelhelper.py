##################################################
## testy.py
##################################################
## Author: clenahan@cloud2gnd.com
## Copyright: Copyright 2024
## Version: 1.0
##################################################

def setHeaderCell(sheet,nCol,text,width):
     sheet.cell(row = 1, column=nCol).value = text
     sheet.column_dimensions[chr(64+nCol)].width = width
