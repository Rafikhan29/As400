import xlrd
import time
import os
from win32com.client import Dispatch
from xlwt import Workbook

class CustomLibrary:

    def __init__(self):
        pass

    def get_latest_file_in_folder(self,folderpath,filestartname='None'):
        fileslist = os.listdir(folderpath)
        screenshotNum = 0
        for fileName in fileslist:
            bStatus = int(str(fileName).find(filestartname))>=0
            if not bStatus:
                continue
            fileNumber = int(str(fileName).replace(".jpg","").split("_")[1])
            if fileNumber > screenshotNum:
                screenshotNum = fileNumber
        filepath = folderpath+"\\screenshot_"+str(screenshotNum)+".jpg"
        print filepath
        if not os.path.exists(filepath):
            return "NA"
        return filepath

    def create_ms_excel_file_using_existing_file(self,inputFilePath,outputFilePath):
        """ It retuen the list of registration codes"""
        book = Workbook()
        workbook = xlrd.open_workbook(inputFilePath)
        snames=workbook.sheet_names()
        expectedColumNumber=-1
        for oldSheetName in snames:
            opworksheet = book.add_sheet(oldSheetName)
            worksheet=workbook.sheet_by_name(oldSheetName)
            noofrows=worksheet.nrows
            tempList=[]
            
            for rowno in range(0,noofrows):
                row=worksheet.row(rowno)
                for colno in range(0,len(row)):
                    cellval=worksheet.cell_value(rowno,colno)
                    if cellval.lower()=='status':
                        expectedColumNumber = colno
                    if colno==expectedColumNumber and rowno >= 1:
                        opworksheet.write(rowno,colno,"Not Executed")
                    else:
                        opworksheet.write(rowno,colno,cellval)
        book.save(outputFilePath)
    
    def updated_ms_excel_file(self,strFilePath,strsheetName,dctVarb):
        """ It retuen the list of registration codes"""
        try:
            exlObj = Dispatch("Excel.Application")
            exlObj.Application.Visible=False
            workbook = exlObj.Workbooks.Open(strFilePath)
            worksheet = workbook.Worksheets(strsheetName)
            colNames=[]
            used = worksheet.UsedRange
            
            intRowsCount =used.Row+used.Rows.Count-1
            #print "intRowsCount: "+str(intRowsCount)
            intColCount =used.Column + used.Columns.Count - 1
            #print "intColCount: "+str(intColCount)
            for iRowIndex in range(1,intRowsCount+1):
              for iColIndex in range(1,intColCount+1):
                cellValue = worksheet.Cells(iRowIndex,iColIndex).Value
                cellValue = str(cellValue)
                if iRowIndex==1:
                  colNames.append(cellValue)
                  continue
                if cellValue!=dctVarb['RecordNumber']:
                  continue
                worksheet.Cells(iRowIndex,int(colNames.index("Status"))+1).Value = dctVarb['Status']
                worksheet.Cells(iRowIndex,int(colNames.index("Message"))+1).Value = dctVarb['Message']
                worksheet.Cells(iRowIndex,int(colNames.index("ScreenShot"))+1).Value = dctVarb['ScreenShot']
            exlObj.ActiveSheet.Columns.AutoFit()
            workbook.Save()
            workbook.close
            exlObj.Application.Quit()
        except Exception as exp:
          print exp
          try:
            workbook.Save()
            workbook.close
            exlObj.Application.Quit()
          except:
            print "exp"
    
    def read_multiple_testdata(self,filepath,sheetname,testcasename):
        """read multiple rows of testdata based on testcase name"""
        try:
            workbook = xlrd.open_workbook(filepath)
            worksheet = workbook.sheet_by_name(sheetname)
            noofrows = worksheet.nrows
            print "noofrows: "+ str(noofrows)
            dictvar={}
            index=1
            for rowno in range(0,noofrows):
                cellvalue = worksheet.cell_value(rowno,0)
                rowValues = worksheet.row_values(rowno)
                if cellvalue == testcasename:
                    tempdict = {}
                    for colno in range(0,len(rowValues)):
                        keydata = worksheet.cell_value(0,colno)
                        celdata = worksheet.cell_value(rowno,colno)
                        if len(str(keydata))==0:
                            continue
                        if len(str(celdata))==0:
                            celdata = ""
                        tempdict[keydata] = celdata
                    dictvar[str(index)] = tempdict
                    index+=1
            return dictvar
        except Exception as exp:
            print "Got exception in read_multiple_testdata keyword.Error: "+str(exp)
            return {}


    def read_all_testdata(self,filepath,sheetname):
        """read multiple rows of testdata based on testcase name"""
        try:
            workbook = xlrd.open_workbook(filepath)
            worksheet = workbook.sheet_by_name(sheetname)
            noofrows = worksheet.nrows
            print "noofrows: "+ str(noofrows)
            dictvar={}
            for rowno in range(1,noofrows):
                rowno = int(rowno)
                cellvalue = worksheet.cell_value(rowno,0)
                rowValues = worksheet.row_values(rowno)
                tempdict = {}
                for colno in range(0,len(rowValues)):
                    keydata = worksheet.cell_value(0,colno)
                    celdata = worksheet.cell_value(rowno,colno)
                    if len(str(keydata))==0:
                        continue
                    if len(str(celdata))==0:
                        celdata = ""
                    tempdict[keydata] = celdata
                dictvar[str(index)] = tempdict 
            return dictvar
        except Exception as exp:
            print "Got exception in read_multiple_testdata keyword.Error: "+str(exp)
            return {}
