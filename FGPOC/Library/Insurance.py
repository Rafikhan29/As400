import time
import wmi
from win32com.client import Dispatch
import os
from robot.libraries.BuiltIn import BuiltIn


class Insurance:

    def __init__(self):
        pass

    def open_console(self,AppConnPath):
        """Invokes the Emulator by passing the .ws file path"""
        try:
            autECLConnMgr.autECLConnList.Refresh()
            if autECLConnMgr.autECLConnList.Count == 0:
                if not os.path.exists(AppConnPath):
                    print "File path doesnt Exists: "+AppConnPath
                    return False
                else:
                    print "File path Exists: "+AppConnPath
                autECLConnMgr.StartConnection("profile=" + AppConnPath + " winstate=MAX") 
                timeout = 0
                while (timeout < 10):
                    processList = wmi.WMI()
                    for process in processList.Win32_Process ():
                        if process.Name.lower()=='pcsws.exe':
                            return True
                    timeout = timeout + 1
                print "Connection timout. Emulator dint open in 10 secs"
                return False
            else:
                self.connection_reset()
                return True
        except Exception as exp:
            print "Got exception in open_console keyword.Error: "+str(exp)
            return False

    def shutdown_console(self):
        """Stops the ongoing connection and closes the Emulator"""
        try:
            autECLConnMgr = Dispatch("PCOMM.autECLConnMgr")
            autECLConnList = Dispatch("PCOMM.autECLConnList")
            autECLPSObj = Dispatch("PCOMM.autECLPS")

            if autECLConnMgr.autECLConnList.Count > 0:
                autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
                autECLPSObj.StopCommunication()
                autECLConnMgr.StopConnection(autECLConnMgr.autECLConnList(1).Handle, "saveprofile=no")
            return True
        except Exception as exp:
            print "Got exception in shutdown_console keyword.Error: "+str(exp)
            return False

    def connection_reset(self):
        """resets the avaiable connection"""
        try:
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")
            
            autECLConnList.Refresh()
            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            return True
        except Exception as exp:
            print "Got exception in connection_reset keyword.Error: "+str(exp)
            return False
        
    def get_connection_count(self):
        """gets the available number of connections"""
        autECLConnMgr = Dispatch("PCOMM.autECLConnMgr")
        return autECLConnMgr.autECLConnList.Count
    
    def wait_for_text(self, sSearchText, iTime=5):
        """Waits for the text for the time mentioned, to be displayed in the screen. Time in seconds"""
        try:
            iTime = int(iTime)
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")

            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            row = 1
            col = 1
            timeout = 0
            while (timeout < iTime):
                autECLPSObj.autECLFieldList.Refresh()
                if autECLPSObj.SearchText(sSearchText, 1, row, col)[0]:
                    return True
                timeout = timeout + 1
                time.sleep(1)
            print "No Text found: "+sSearchText
            self.capture_screenshot()
            return False
        except Exception as exp:
            print "Got exception in wait_for_text keyword.Error: "+str(exp)
            self.capture_screenshot()
            return False

    def wait_for_text_until_invisible(self, sSearchText, iTime=10):
        """Waits for the text until it is invisible on screen. Time in seconds"""
        try:
            iTime = int(iTime)
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")

            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            row = 1
            col = 1
            timeout = 0
            while (timeout < iTime):
                autECLPSObj.autECLFieldList.Refresh()
                if not autECLPSObj.SearchText(sSearchText, 1, row, col)[0]:
                    return True
                timeout = timeout + 1
                time.sleep(1)
            print "Text is still visible"
            self.capture_screenshot()
            return False
        except Exception as exp:
            print "Got exception in wait_for_text_until_invisible keyword.Error: "+str(exp)
            return False
        

    def press_key(self, Keyvalue,count=1):
        """"Use to perform keyboard events."""
        try:
            count = int(count)
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")
            
            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            autECLPSObj.autECLFieldList.Refresh()
            if count == 1:
                autECLPSObj.SendKeys(Keyvalue)
                return True
            
            for i in range(0,count):
                autECLPSObj.SendKeys(Keyvalue)
                time.sleep(2)
            return True
        except Exception as exp:
            print "Got exception in press_key.Error: "+str(exp)
            return False

    def capture_screenshot(self):
        """"It will capture the screenshots based on ${globalScreenShot}  global variable value."""
        try:
            screenshot = BuiltIn().get_library_instance("Screenshot")
            bStatus = BuiltIn().get_variable_value("${globalScreenShot}")
            print "bStatus: "+str(bStatus)
            if(str(bStatus).lower()=='true'):
                screenshot.take_screenshot()
        except Exception as exp:
            print "Got exception in capture_screenshot keyword.Error: "+str(exp)
            return False
        
    def get_value_by_field_name(self, sFieldName, iPos=0):
        """To capture the Field Text by the given Field Label """
        try:
            iPos = int(iPos)
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")
            
            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            autECLPSObj.autECLFieldList.Refresh()
            i = 1
            row = 1
            col = 1
            if not autECLPSObj.SearchText(sFieldName, 1, row, col)[0]:
                return "NA"
            while i <= autECLPSObj.autECLFieldList.Count:
                if autECLPSObj.autECLFieldList(i).GetText().strip() == sFieldName:
                    return autECLPSObj.autECLFieldList(i+iPos).GetText()
                i = i + 1
        except Exception as ex:
            print "Got Exception in get_value_by_field_nameL.Error: "+str(ex)
            return "NA"

    def extract_all_fields_from_screen(self, sFilePath=None):
        """Extract all the field names to a given text file with their Field Positions"""
        try:
            if sFilePath == None:
                sFilePath = "FieldDetails.txt"
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")

            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            autECLPSObj.autECLFieldList.Refresh()
            if autECLPSObj.autECLFieldList.Count==0:
                self.write_text_file("*************No Fields Available***************", sFilePath, False)
                print "No fields available on the screen"
                return False
            self.write_text_file("*************Extracting all the fields in the Screen***************", sFilePath, False)
            index = 1
            print "Field List.Count:"
            print autECLPSObj.autECLFieldList.Count
            while index < autECLPSObj.autECLFieldList.Count:
                self.write_text_file("Field Index: "+str(index)+" Field label: " + autECLPSObj.autECLFieldList(index).GetText(), sFilePath, True)
                index = index + 1
            self.write_text_file("*************End of - Extracting all the fields in the Screen***************", sFilePath, True)
            return True
        except Exception as exp:
            print "Got exception in extract_all_fields_from_screen keyword.Error: "+str(exp)
            return False

    def write_text_file(self, sText, fPath, append=True):
        """write data to given text file"""
        if append:
            myFile = open(fPath, 'a')
        else:
             myFile = open(fPath, 'w')
        myFile.write(str(sText)+"\n")
        myFile.close()
        return True

    def enter_text_by_field_Name(self, sSearchText, sValue="", instance=1):
        """To enter the text against a field label. The first parameter is mandate and other two are optional.Instance is used in the case of multiple fields with same name.If you want to erase the field value, pass Value as "empty" from your test data."""
        try:
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")

            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            autECLPSObj.autECLFieldList.Refresh()
            if sSearchText != "":
                if str(sValue).lower()=='na':
                    return True
                bstatus = self.set_cursor_position(sSearchText,instance)
                if not bstatus:
                    return False
                autECLPSObj.SendKeys("[TAB]")
                time.sleep(1)
                autECLPSObj.SendKeys("[erase eof]")
                if sValue != "":
                    autECLPSObj.setText(sValue)
                    return True
                elif sValue == "":
                    print "No value in Sheet"
                    return True
                else:
                    return False
            else:
                return False
        except Exception as exp:
            
            print "Got Exception in enter_text_by_field_Name.Error: "+str(exp)
            self.capture_screenshot()
            return False

    def set_cursor_position(self, sSearchText, instance=1):
        """To set the cursor at the start of the text passed as parameter.Instance is used if we have more than 1 identical text in the screen."""
        try:
            instance = int(instance)
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")
            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            
            autECLPSObj.autECLFieldList.Refresh()
            row = 1
            col = 1
            newcol = 0
            Temprow = 0
            temp = autECLPSObj.SearchText(sSearchText, 1, row, col)
            if temp[0]:
                row = temp[1]
                col = temp[2]
                if instance > 1:
                    ints = 1
                    while ints <= instance:
                        newcol = newcol+ 1
                        autECLPSObj.SetCursorPos(row, newcol)
                        result = autECLPSObj.SearchText(sSearchText, 1, row, newcol)
                        if result[0]:
                            row = result[1]
                            newcol = result[2]
                            if (ints == instance):
                                autECLPSObj.SetCursorPos(row, newcol)
                                return True
                            ints = ints + 1
                else:
                    autECLPSObj.SetCursorPos(row, col)
                    return True
            else:
                self.capture_screenshot()
                return False
            
        except Exception as exp:
            print "Got Exception in set_cursor_position keyword. Error: "+str(exp)
            self.capture_screenshot()
            return False

    def set_cursor_position_for_menu(self, sSearchMenuText, instance=1):
        """To set the cursor at the start of the menutext passed as parameter.Instance is used if we have more than 1 identical text in the screen."""
        try:
            instance = int(instance)
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")
            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            
            autECLPSObj.autECLFieldList.Refresh()
            row = 1
            col = 1
            newcol = 0
            Temprow = 0
            temp = autECLPSObj.SearchText(sSearchMenuText, 1, row, col)
            if temp[0]:
                row = temp[1]
                col = temp[2]
                if instance > 1:
                    ints = 1
                    while ints <= instance:
                        newcol = newcol+ 1
                        autECLPSObj.SetCursorPos(row, newcol)
                        result = autECLPSObj.SearchText(sSearchMenuText, 1, row, newcol)
                        if result[0]:
                            row = result[1]
                            newcol = result[2]
                            if (ints == instance):
                                autECLPSObj.SetCursorPos(row, newcol - 4)
                                return True
                            ints = ints + 1
                else:
                    autECLPSObj.SetCursorPos(row, col - 4)
                    return True
            else:
                print sSearchMenuText+" no available on screen"
                self.capture_screenshot()
                return False
            
        except Exception as exp:
            self.capture_screenshot()
            print "Got Exception in set_cursor_position_for_menu keyword. Error: "+str(exp)
            return False

    def select_menu_Item(self, sMenuName, instance=1):
        """To Select the Module or Sub module."""
        try:
            bstatus = self.set_cursor_position_for_menu(sMenuName,instance)
            if not bstatus:
                return False
            self.press_key("[TAB]")
            self.press_key("[ENTER]")
            return True
        except Exception as exp:
            print "Got Exception in select_menu_Item keyword. Error: "+str(exp)
            return False
            

    def get_cursor_position(self):
        """This gets the current position of the cursor in the presentation space for the connection associated"""
        try:
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")
            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            
            autECLPSObj.autECLFieldList.Refresh()
            curRow = autECLPSObj.CursorPosRow
            curCol = autECLPSObj.CursorPosCol
            return (curRow,curCol)
        except Exception as exp:
            print "Got Exception in get_cursor_position keyword. Error: "+str(exp)
            return (0,0)

    def check_and_mark(self,sSearchVal,keyOpr="[BackTab]",markVal="X"):
        """This gets the current position of the cursor in the presentation space for the connection associated"""
        try:
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")

            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            autECLPSObj.autECLFieldList.Refresh()
            if not (self.set_cursor_position(sSearchVal)):
                return False
            self.press_key(keyOpr)
            self.enter_text(markVal)
            return (curRow,curCol)
        except Exception as exp:
            print "Got Exception in check_and_mark keyword. Error: "+str(exp)
            return (0,0)

    def enter_text(self, sValue):
        """ This keyword will enter text svalue at current position"""
        try:
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")

            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            autECLPSObj.autECLFieldList.Refresh()
            autECLPSObj.setText(sValue)
            return True
        except Exception as exp:
            print "Got Exception in enter_text keyword. Error: "+str(exp)
            return False

    def get_value_by_row_and_column(self, Row, Col, txtLen):
        """To capture the output values or any other values on the screen based on the row, column and length of the text"""
        try:
            Row = int(Row)
            Col = int(Col)
            txtLen = int(txtLen)
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")
            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            return autECLPSObj.GetText(Row, Col, txtLen)
        except Exception as exp:
            print "Got Exception in get_value_by_row_and_column keyword. Error: "+str(exp)
            return "NA"

    def select_item_from_search_table_by_field_name(self,fieldname,selectvalue,instance=1):
        """To capture the output values or any other values on the screen based on the row, column and length of the text"""
        try:
            autECLPSObj = Dispatch("PCOMM.autECLPS")
            autECLConnList = Dispatch("PCOMM.autECLConnList")

            autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
            autECLPSObj.autECLFieldList.Refresh()
            if str(selectvalue).lower()=='na':
                return True
            fieldStatus = self.wait_for_text(fieldname,10)
            bstatus = self.set_cursor_position(fieldname,instance)
            if not bstatus:
                return False
            time.sleep(1)
            self.press_key("[TAB]")
            time.sleep(1)
            self.press_key("[erase eof]")
            self.press_key("[PF4]")
            bstatus = self.wait_for_text("Table Item Search",10)
            if not bstatus:
                return False
            for index in range(1,20):
                bstatus = self.set_cursor_position(selectvalue,1)
                if not bstatus:
                    bMoreStatus = self.wait_for_text("More...",3)
                    if not bMoreStatus:
                        self.press_key("[ENTER]")
                        return False
                    else:
                        self.press_key("[pagedn]")
                else:
                    self.press_key("[backtab]")
                    self.press_key("1")
                    self.press_key("[ENTER]")
                    tableStatus = self.wait_for_text_until_invisible("Table Item Search",10)
                    return tableStatus
            self.press_key("[ENTER]")
            return False
        except Exception as exp:
            print "Got Exception in select_item_from_search_table_by_field_name keyword. Error: "+str(exp)
            return "NA"
           
# Check Below Keywords *********************************************************************
#*******************************************************************************************

    def go_to_screen(self, sScreenName, KeyValue='[PF3]'):
        """Will perform the keyboard operation until the text/screen is displayed"""
        autECLPSObj = Dispatch("PCOMM.autECLPS")
        autECLConnList = Dispatch("PCOMM.autECLConnList")
        autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
        
        KeyCnt = 0
        row = 1
        col = 1
        autECLPSObj.autECLFieldList.Refresh()
        if autECLPSObj.SearchText(sScreenName, 1, row, col)[0]:
            return True
        KeyCnt = 0
        while KeyCnt <= 20:
            print "press"+str(KeyCnt)
            autECLPSObj.SendKeys(KeyValue)
            autECLPSObj.autECLFieldList.Refresh()
            if self.wait_for_text(sScreenName,5):
                return True
            KeyCnt = KeyCnt + 1
        return False





    def get_value_by_rectangle(self, StartRow, StartCol, EndRow, EndCol):
        """To get all the data present in the rectangle formed by the StartRow, StartCol, EndRow  and EndCol."""
        autECLPSObj = Dispatch("PCOMM.autECLPS")
        autECLConnList = Dispatch("PCOMM.autECLConnList")
        autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
        
        return autECLPSObj.GetTextRect(StartRow, StartCol, EndRow, EndCol)

    def set_cursor_position_in_backward_direction(self, sSearchText, instance=1):
        """To set the cursor before the text from the bottom of the screen in backward direction passed as parameter."""
        autECLPSObj = Dispatch("PCOMM.autECLPS")
        autECLConnList = Dispatch("PCOMM.autECLConnList")
        autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
        
        autECLPSObj.autECLFieldList.Refresh()
        row = 24
        col = 1
        newcol = 80
        Temprow = 0
        temp = autECLPSObj.SearchText(sSearchText, 2, row, col)
        if temp[0]:
            row = temp[1]
            col = temp[2]
            if instance > 1:
                ints = 1
                while (ints <= instance):
                    newcol = newcol - 1
                    autECLPSObj.SetCursorPos(row, newcol)
                    result = autECLPSObj.SearchText(sSearchText, 2, row, newcol)
                    if result[0]:
                        row = result[1]
                        newcol = result[2]
                        if (ints == instance and Temprow > row):
                            autECLPSObj.SetCursorPos(row, newcol - 1)
                            return True
                        Temprow = row
                        ints = ints + 1
            else:
                autECLPSObj.SetCursorPos(row, col - 1)
                return True
        else:
            return False

    def set_cursor_position_before_value(self, sSearchText, instance=1):
        """To set the cursor before the text passed as parameter.Instance is used if we have more than 1 identical text in the screen."""
        autECLPSObj = Dispatch("PCOMM.autECLPS")
        autECLConnList = Dispatch("PCOMM.autECLConnList")
        autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
        
        autECLPSObj.autECLFieldList.Refresh()
        row = 1
        col = 1
        newcol = 0
        temp = autECLPSObj.SearchText(sSearchText, 1, row, col)
        if temp[0]:
            row = temp[1]
            col = temp[2]
            if instance > 1:
                ints = 1
                while (ints <= instance):
                    newcol = col + 1
                    autECLPSObj.SetCursorPos(row, newcol)
                    result = autECLPSObj.SearchText(sSearchText, 1, row, newcol)
                    if result[0]:
                        row = result[1]
                        newcol =  result[2]
                        if (ints == instance):
                            autECLPSObj.SetCursorPos(row, newcol-1)
                            return True
                        ints = ints+1
            else:
                autECLPSObj.SetCursorPos(row, col-1)
                return True
        else:
            return False

    def set_cursor_position_dup(self, sSearchText, row, col, instance=1):
        """Searches the given text from the given row and column and sets the cursor at the given text."""
        autECLPSObj = Dispatch("PCOMM.autECLPS")
        autECLConnList = Dispatch("PCOMM.autECLConnList")
        autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
        
        autECLPSObj.autECLFieldList.Refresh()
        newcol = 0
        temp = autECLPSObj.SearchText(sSearchText, 1, row, col)
        if temp[0]:
            row = temp[1]
            col = temp[2]
            if instance > 1:
                newcol = col + 1
                autECLPSObj.SetCursorPos(row, newcol)
                result = self._autECLPSObj.SearchText(sSearchText, 1, row, newcol)
                if result[0]:
                    row = result[1]
                    newcol = result[2]
                    autECLPSObj.SetCursorPos(row, newcol - 1)
                    return True
            else:
                autECLPSObj.SetCursorPos(row, col - 1)
                return True
        else:
            return False





    def edit_and_update_value(self, sSearchText, row, col, sValue):
        """Update the value of a editable field with new value provided."""
        autECLPSObj = Dispatch("PCOMM.autECLPS")
        autECLConnList = Dispatch("PCOMM.autECLConnList")
        autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
        
        bstatus = set_cursor_position_dup(sSearchText, row, col)
        if not bstatus:
            return False
        press_key("[Tab]")
        enter_text(sValue)
        return True


    def validate_text_on_screen(self, sSearchText, instance=1):
        autECLPSObj = Dispatch("PCOMM.autECLPS")
        autECLConnList = Dispatch("PCOMM.autECLConnList")
        autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
        
        autECLPSObj.autECLFieldList.Refresh()
        row = 1
        col = 1
        newcol = 1
        Temprow = 0
        if instance > 1:
            ints = 1
            while ints <= instance:
                newcol = newcol + 1
                autECLPSObj.SetCursorPos(row, newcol)
                temp = autECLPSObj.SearchText(sSearchText, 1, row, newcol)
                if temp[0]:
                    row= temp[1]
                    newcol = temp[2]
                    if (ints == instance):
                        return True
                    ints = ints + 1
        else:
            if autECLPSObj.SearchText(sSearchText, row, col)[0]:
                return True
            else:
                return False

    def enter_text_by_field_name_back(self, sSearchText, sValue, instance=1):
        """To enter the text against a field label (ex: if the text field is before the field label). The first parameter is mandate and other two are optional.
        Instance is used in the case of multiple fields with same name"""
        autECLPSObj = Dispatch("PCOMM.autECLPS")
        autECLConnList = Dispatch("PCOMM.autECLConnList")
        autECLPSObj.SetConnectionByHandle(autECLConnList(1).Handle)
        
        bstatus = set_cursor_position(sSearchText,instance)
        if not bstatus:
            return False
        autECLPSObj.SendKeys("[backtab]")
        autECLPSObj.setText(sValue)
        return True


    

    




    


    
        
        

        
    
