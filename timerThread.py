from PyQt5.QtCore import QObject, QThread, pyqtSignal, pyqtSlot

import sys
import visa
import struct
from time import sleep
import time
import chromaLib
import loadFunctions_Chroma_63600
import xlsxwriter
from xlsxwriter.utility import xl_range
from datetime import datetime

version_Info = "Python PyQt5 Slammer Frequency Sweep V0.1"

class TimerThread(QObject):

    testList = {}
    dutList = {}
    step_Count = None
    startStop = True
    str_load_frame_id = None
    load = None
     
    timerSignal = pyqtSignal(int)
    quitSignal = pyqtSignal(object)
    testInfoSignal = pyqtSignal(str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str, str)
   
    excel_command_new_file = "file" 
    excel_command_new_chart_sheet = "sheet" 
    excel_command_close = "close"

    max_samples = 4096 # Chroma 63600 max = 4096 
    min_sample_period = 2e-6 # Chroma 63600 min = 2us 
    max_sample_period = 0.040 # Chroma 63600 max = 40ms 
    sample_period_resolution = 2e-6 # Chroma 63600 sample time resolution = 2us 
    min_duration = max_samples * min_sample_period # assumes always acquiring max_samples 
    max_duration = max_samples * max_sample_period 
    
    sample_Period = None
    rawDataSheet = None
    raw_sheet_row = None
    str_chart_sheet_name = None
    chart_sheet_row = None 
    chart_sheet_col = None 
    headerCellFormat = None
    chart_sheet_number = None
    
    fileName = ""
    testID1 = ""
    testID2 = ""
    testID3 = ""
    portName1 = ""
    portName2 = ""
    portName3 = ""
    amps1 = ""
    amps2 = ""
    amps3 = ""
    duration1 = ""
    duration2 = ""
    duration3 = ""
    trigger1 = ""
    trigger2 = ""
    trigger3 = ""
    pulseTime1 = ""
    pulseTime2 = ""
    pulseTime3 = ""
    
    
    

    def __init__(self, load, str_load_frame_id, step_Count, dutList, testList):
        QThread.__init__(self)
        
        self.testList = testList
        self.dutList = dutList
        self.step_Count = step_Count
        self.str_load_frame_id = str_load_frame_id
        self.load = load
        
    def __del__(self):
        print("del")
#        self.wait()
            
    def startTimer(self):
        self.startStop = True

    def stopTimer(self):
        self.startStop = False
        self.quitSignal.emit(self.quitSignal)
        #This sleep is here so thread doesn't exit before being killed by main routine
        #This is my hyposis about why the quitLoop in main gives continuous errors??
        time.sleep(1)
        
    def testEnd(self):
                    
        print("Test Finished")
        self.stopTimer()
        
    def emitTestInfoSignal(self):
        self.testInfoSignal.emit(self.fileName, str(self.testID1), str(self.testID2), str(self.testID3), self.portName1, self.portName2, 
                                 self.portName3, str(self.amps1), str(self.amps2), str(self.amps3), str(self.duration1), 
                                 str(self.duration2), str(self.duration3), str(self.trigger1), str(self.trigger2), str(self.trigger3),
                                 str(self.pulseTime1), str(self.pulseTime2), str(self.pulseTime3)) 

    def run(self):
        step_loop = 1
        result_loop = 1
        thisTest = {}
        All_Results = {}
        All_Stats = {}
        TestGroup = {}
        this_test_group = 0
        
        self.fileName = self.testList[0]['file_Name']
        self.emitTestInfoSignal()
                    
        while self.startStop:
            print('\n')
            #testList contains entries for every line of the config file
            #The number of primary keys in testList is the number of lines on the sequence tab
            #step_Loop loops through each key in testList
            
            #Get local variables from each row/key
            thisTest = self.testList[step_loop]
            test_Id = thisTest['TestID']
            console = thisTest['Console']
            port = thisTest['Port']
            frame = thisTest['Frame']
            channel = thisTest['Channel']
            amps = thisTest['Amps']
            duration = thisTest['Duration'] 
            trigger = thisTest['Trigger'] 
            excel_command = thisTest['ExcelCommand']
            rise_slew_rate = thisTest['RiseSlewRate'] 
            fall_slew_rate = thisTest['FallSlewRate'] 
			
            dut_config = self.dutList[console]['dut_config'] 
            dut_productSn = self.dutList[console]['dut_productSn'] 
            dut_tag = self.dutList[console]['dut_tag'] 

            test_group_size = 1 

            #Configuration spreadsheets need to have 'file' on the first line.
            #Not sure if multiple files works, this is legacy stuf from Davids code.
            if excel_command == self.excel_command_new_file: # new workbook file 
                pending_excel_command = excel_command 
                print("Opening new workbook for next test.") 
            elif (excel_command == self.excel_command_new_chart_sheet) and (pending_excel_command != self.excel_command_new_file) : # new chart sheet 
                #excel command could be null
                pending_excel_command = excel_command 
                print("Creating new chart sheet for next test.")
            
            			
            if int(test_Id / 10) == 111: # 1110 to 1119 ("Constant-Current / 1 Port / Short Duration") 
                print("Testing '" + str(dut_tag) + "' port '" + str(port) + "' on frame-load number " + str(frame) + "-" + str(channel))
                #For this test only one port is loaded at a time.
                print("Single Port Load")
                #All_Results is contains a single set of V and I data.
                All_Results[result_loop], self.sample_Period = loadFunctions_Chroma_63600.constantCurrentTest2(self, self.str_load_frame_id, port, channel, amps, duration, trigger, rise_slew_rate, fall_slew_rate) 
            
            elif int(test_Id / 10) == 131: # 1110 to 1119 ("Constant-Current / 1 Port / Short Duration") 
                print("Testing '" + str(dut_tag) + "' port '" + str(port) + "' on frame-load number " + str(frame) + "-" + str(channel))
                #For this test only one port is loaded at a time.
                print("Single Port HDMI Ramp Load")
                #All_Results is contains a single set of V and I data.
                All_Results[result_loop], self.sample_Period = loadFunctions_Chroma_63600.singlePortRamp(self, self.str_load_frame_id, port, channel, amps, duration, trigger) 
            
            elif int(test_Id / 10) == 121: #Multiple port test
                print("Group Port Load")
                test_group_index = 0
                this_test_group = this_test_group + 1
                
                TestGroup[0] = this_test_group
                
                #Find all lines with same test ID between excel new sheet commands
                #It's the new sheet command that determines a test group
                for step_Loop2 in range (step_loop, self.step_Count):
                    ThisTest2 = self.testList[step_Loop2]
                    test_id2 = ThisTest2['TestID']
                    console2 = ThisTest2['Console']
                    port2 = ThisTest2['Port']
                    excel_command2 = ThisTest2['ExcelCommand']
                    
                    #Put first test in group, group index is zero
                    #Only put following tests into group if they have the same test ID and do not have an excel command.
                    if test_group_index == 0 or (test_id2 == test_Id and excel_command2 !=  self.excel_command_new_file and excel_command2 != self.excel_command_new_chart_sheet):
                        test_group_index = test_group_index + 1
                        #Add test to test group
                        #Basically take a the next line in the sequence tab and add it to this new group of tests
                        print "test added to group " + port2
                        #Add keys to TestGroup. Make each a row from the spreadsheet
                        TestGroup[test_group_index] = ThisTest2
                        #These values are only used to send display info back to main window. Only groups of three supported now.
                        if test_group_index == 1:
                            self.testID1 = ThisTest2['TestID']
                            self.portName1 = ThisTest2['Port']
                            self.amps1 = ThisTest2['Amps']
                            self.duration1 = ThisTest2['Duration']
                            self.trigger1 = ThisTest2['Trigger']
                            self.pulseTime1 = ThisTest2['PulseT2']
                        elif test_group_index == 2:
                            self.testID2 = ThisTest2['TestID']
                            self.portName2 = ThisTest2['Port']
                            self.amps2 = ThisTest2['Amps']
                            self.duration2 = ThisTest2['Duration']
                            self.trigger2 = ThisTest2['Trigger']
                            self.pulseTime2 = ThisTest2['PulseT2']
                        elif test_group_index == 3:
                            self.testID3 = ThisTest2['TestID']
                            self.portName3 = ThisTest2['Port']
                            self.amps3 = ThisTest2['Amps']
                            self.duration3 = ThisTest2['Duration']
                            self.trigger3 = ThisTest2['Trigger']
                            self.pulseTime3 = ThisTest2['PulseT2']
                            
                            self.emitTestInfoSignal() 
                    else:
                        #Exit for step_loop2 when first different test ID found or new sheet command is found
                        break
                #End for step_Loop2
                
                test_group_size = test_group_index
                test_group_index = 0
                if test_group_size < 2:
                    print "ERROR: Several identical Test IDs expected for simultaneous multiple port test! Skipping test."
                else:
                    GroupResults = {}
                    print "Starting Group Test"
#                    print(TestGroup)
#                    for group in TestGroup:
#                        print(group)
#                        print(TestGroup[group])
        
                    #TestGroup is a list of the size of the number of tests grouped, usually 3. Each one is an instance or row from the config spreadsheet.
                    GroupResults, self.sample_Period = loadFunctions_Chroma_63600.constantCurrentTestMultPulse(self, self.str_load_frame_id, TestGroup)
                    for result_loop2 in range (result_loop, result_loop + test_group_size):
                        test_group_index = test_group_index + 1
                        #All_Results grows here. We add keys to match the size of the group
                        All_Results[result_loop2] = GroupResults[test_group_index]
                        
            elif int(test_Id / 10) == 313:
                print("Testing Ramp")
                test_group_index = 0
                this_test_group = this_test_group + 1
                
                TestGroup[0] = this_test_group
                
                #Find all lines with same test ID between excel new sheet commands
                #It's the new sheet command that determines a test group
                for step_Loop2 in range (step_loop, self.step_Count):
                    ThisTest2 = self.testList[step_Loop2]
                    test_id2 = ThisTest2['TestID']
                    console2 = ThisTest2['Console']
                    port2 = ThisTest2['Port']
                    excel_command2 = ThisTest2['ExcelCommand']
                    
                    #Put first test in group, group index is zero
                    #Only put following tests into group if they have the same test ID and do not have an excel command.
                    if test_group_index == 0 or (test_id2 == test_Id and excel_command2 !=  self.excel_command_new_file and excel_command2 != self.excel_command_new_chart_sheet):
                        test_group_index = test_group_index + 1
                        #Add test to test group
                        #Basically take a the next line in the sequence tab and add it to this new group of tests
                        print "test added to group " + port2
                        TestGroup[test_group_index] = ThisTest2
                        if test_group_index == 1:
                            self.testID1 = ThisTest2['TestID']
                            self.portName1 = ThisTest2['Port']
                            self.amps1 = ThisTest2['Amps']
                            self.duration1 = ThisTest2['Duration']
                            self.trigger1 = ThisTest2['Trigger']
                            self.pulseTime1 = ThisTest2['PulseT2']
                        elif test_group_index == 2:
                            self.testID2 = ThisTest2['TestID']
                            self.portName2 = ThisTest2['Port']
                            self.amps2 = ThisTest2['Amps']
                            self.duration2 = ThisTest2['Duration']
                            self.trigger2 = ThisTest2['Trigger']
                            self.pulseTime2 = ThisTest2['PulseT2']
                        elif test_group_index == 3:
                            self.testID3 = ThisTest2['TestID']
                            self.portName3 = ThisTest2['Port']
                            self.amps3 = ThisTest2['Amps']
                            self.duration3 = ThisTest2['Duration']
                            self.trigger3 = ThisTest2['Trigger']
                            self.pulseTime3 = ThisTest2['PulseT2']
                            
                            self.emitTestInfoSignal() 
                    else:
                        #Exit for step_loop2 when first different test ID found or new sheet command is found
                        break
                #End for step_Loop2
                
                test_group_size = test_group_index
                test_group_index = 0
                if test_group_size < 2:
                    print "ERROR: Several identical Test IDs expected for simultaneous multiple port test! Skipping test."
                else:
                    GroupResults = {}
                    print "Starting Group Test"
#                    print(TestGroup)
#                    for group in TestGroup:
#                        print(group)
#                        print(TestGroup[group])
        
                    GroupResults, self.sample_Period = loadFunctions_Chroma_63600.thresholdTest(self, self.str_load_frame_id, TestGroup)
                    for result_loop2 in range (result_loop, result_loop + test_group_size):
                        test_group_index = test_group_index + 1
                        All_Results[result_loop2] = GroupResults[test_group_index]
                
 
            #End if test ID 121
            #test_group_size is 1 for 1 port constant current so in that case this gets done once.
            #step_loop is incremented here which takes us to the next row or skips rows for group tests
            for test_group_index in range (1, test_group_size + 1):
                All_Results[result_loop]['test_step'] = step_loop 
                All_Results[result_loop]['dut_config'] = dut_config 
                All_Results[result_loop]['dut_productSn'] = dut_productSn 
                All_Results[result_loop]['dut_tag'] = dut_tag
                
                ThisTest2 = self.testList[step_loop] 
                test_id2 = ThisTest2['TestID'] 
                console2 = ThisTest2['Console']
                dut_productSn = self.dutList[console2]['dut_productSn']  
    				
                # below unnecessary? 
                port2 = ThisTest2['Port']
                excel_command2 = ThisTest2['ExcelCommand'] 
                
#Write CSV stuff was here
                #This keeps getting set in group test as len(All_Reults) increments
                result_count2 = len(All_Results)		
                print "result count " + str(result_count2)
                print "step loop " + str(step_loop)
                All_Results[result_loop]['excel_command'] = pending_excel_command 
                pending_excel_command = 0  
            
                result_loop = result_loop + 1
                step_loop = step_loop + 1
           
            #This gets done once for 1 port, multiples for group test.
            for result_loop2 in range (1, result_count2 + 1):
                print "got to OCP section"
                Raw_Data = All_Results[result_loop2]
                All_Stats[result_loop2] = self.calcOCPData(Raw_Data, port)
                Raw_Data = All_Results[result_loop2]
                Stats = All_Stats[result_loop2]
                self.writeExcelResults(Raw_Data, Stats, self.testList[0])
            
            # Clear results tables
            #We reset All_Results because it is only used once with a single key for one port test or multiple keys for group tests
            result_loop = 1
            All_Results = {}
            All_Stats = {}

            self.timerSignal.emit(step_loop)
            if step_loop >= self.step_Count:
                self.testEnd()
            time.sleep(3)
            
        print ("Run Loop Finished")
        self.workBook.close()
            
           
    def writeExcelResults(self, T_Data, T_Stats, T_TestSeqInfo):

        DataTableTags = ["test_step", "Test Step #", True, 
            "dut_config", "DUT Config", False, "dut_productSn", "DUT SN", False, "dut_tag", "DUT Tag", False, "dut_port", "DUT Port", False, 
            "load_on_seconds", "Load On (sec)", True, 
            "date_time", "Date Time", False, "load_frame", "Load Frame", False, "load_model", "Load Model", False, "load_channel", "Load Channel", True, 
            "load_set_cc", "Load Set CC", True, "sample_count", "Sample Count", True, "sample_period", "Sample Period", True, 
            "test_code", "Test Code", True, "no_load_volt_post", "Post-test No-load Voltage", True ] # fields in raw data 
	
        datatabletags_field_count = 3

        DataTableTags2 = ["file_folder", "File Folder", False, "file_name", "File Name", False, "file_extension", 
                          "Ext", False, "excel_command", "Excel Command", True]

        DataCodeValues = {}  
        DataCodeValues['test_code'] = {} 
        DataCodeValues['test_code'][0] = "Test Type" # text meaning of code  -- (alternative: for DataTableTags?? "test_type", "Test Type", false) 
        # test code guide - [abcd]: a = type, b = 1 or multiple ports, c = NA duration or short duration (up to continous digitizing time limit, e.g. 163 sec) or long duration or either, d = future use
        DataCodeValues['test_code'][1110] = "Constant-Current / 1 Port / Short Duration" 
        DataCodeValues['test_code'][1310] = "Ramp / 1 Port / Short Duration" 
        DataCodeValues['test_code'][1120] = "Constant-Current / 1 Port / Long Duration" 
        DataCodeValues['test_code'][1210] = "Constant-Current / Multiple Ports / Short Duration" 
        DataCodeValues['test_code'][1220] = "Constant-Current / Multiple Ports / Long Duration" 
        DataCodeValues['test_code'][2100] = "Open Voltage / 1 Port" 
        DataCodeValues['test_code'][3130] = "OCP Ramp / 1 Port / Any Duration"  
        
        # constants (here for now) 
        str_raw_sheet_name = "Raw_Data" 
        str_chart_sheet_name_prefix = "Tests"
        str_chart_sheet_name_suffix = "0"
#        str_chart_sheet_name = "Chart Sheet" 
        color_curr = '#0088f0' # hex color format is BBGGRR 
        color_volt = '#f08800' # hex color format is BBGGRR 
        ppi = 72 # points per inch 
        chart_spec = {} 
        chart_spec['left'] = 0
        chart_spec['top'] = 0
        chart_spec['width'] = ppi * 9.0
        chart_spec['height'] = ppi * 4.0 
        
        new_excel = False
        new_chart_sheet = False 
        
        file_command = T_Data['excel_command']
        print "File Command " + str(file_command)
        
        #new excel file, open workbook and add raw data sheet and first chart sheet
        if file_command == 'file':
            print "New File Started"
            new_excel = True
            new_chart_sheet = True
            self.raw_sheet_row = 1
            raw_sheet_col = 1
            
            excelFileName = "powerPortsTestFile.xlsx"
            self.workBook = xlsxwriter.Workbook(excelFileName)
            self.rawDataSheet = self.workBook.add_worksheet(str_raw_sheet_name)
                
            self.headerCellFormat = self.workBook.add_format()
            self.headerCellFormat.set_font_size(14)
            self.headerCellFormat.set_bold()
            
            print "Opened New Excel Results File"
            
        else:
            print "Adding results to open Excel File"
            new_excel = False
    
        if file_command == 'sheet':
            new_chart_sheet = True
            print "Starting new Excel chart sheet."

        raw_sheet_col = 0
        value_name_col_number = raw_sheet_col # first column number for sample data header, next column is first sample 
        dtt_col_number = raw_sheet_col + 1 # values for items in DataTableTags{} 
        
        #If this is a new file populate raw data sheet with test info
#        if new_excel:
#            seq_file_location = T_TestSeqInfo['file_Location'] 
#            seq_file_name_ext = T_TestSeqInfo['file_Name'] + "." + T_TestSeqInfo['file_Extension'] 
#            self.rawDataSheet.write(self.raw_sheet_row, value_name_col_number, "Sequence File Location") 
#            self.rawDataSheet.write(self.raw_sheet_row, dtt_col_number, seq_file_location)
#            self.raw_sheet_row = self.raw_sheet_row + 1 
#            self.rawDataSheet.write(self.raw_sheet_row , value_name_col_number, "Sequence File Name.Ext" )
#            self.rawDataSheet.write(self.raw_sheet_row , dtt_col_number, seq_file_name_ext)  
#            self.raw_sheet_row = self.raw_sheet_row + 2 # leave blank row 

        member = 0
        value = 0  
        TableValues = 0     
        source_folder = False
        source_name = False
        source_extension = False

        for loop in range(0, len(DataTableTags2), datatabletags_field_count):  # write raw data individual members 
            member = DataTableTags[loop] 
            value = T_Data[member] 
            if value != 0: 
                if member == 'file_folder':
                    source_folder = value
                    source_folder_tag = DataTableTags2[loop + 1]
                elif member == "file_name":
                    source_name = value
                    source_name_tag = DataTableTags2[loop + 1]
                elif member == "file_extension":
                    source_extension = value
                    source_extension_tag = DataTableTags2[loop + 1]
                    
        if source_folder != False:
            self.rawDataSheet.write(self.raw_sheet_row , value_name_col_number, source_folder_tag)
            self.rawDataSheet.write(self.raw_sheet_row, dtt_col_number, source_folder)
            self.raw_sheet_row = self.raw_sheet_row + 1
            
        if source_name != False:
            self.rawDataSheet.write(self.raw_sheet_row , value_name_col_number, source_name_tag)
            self.rawDataSheet.write(self.raw_sheet_row, dtt_col_number, source_name)
            self.raw_sheet_row = self.raw_sheet_row + 1
            
        member = 0
        value = 0  
        test_group1 = 0
        no_load_volt_post_text1 = 0
        #Go through data table tags and put values into spreadsheet
        for loop in range(0, len(DataTableTags), datatabletags_field_count):  # write raw data individual members 
            member = DataTableTags[loop] 
            value = T_Data[member] 
            table_name = 0
            table_value = 0
            if value != 0: 
                text = DataTableTags[loop + 1] 
                is_number = DataTableTags[loop + 2] 
                self.rawDataSheet.write(self.raw_sheet_row , value_name_col_number, text)
                self.rawDataSheet.write(self.raw_sheet_row , dtt_col_number, value) 
			
                if member == 'test_code':
                    # For any coded values, write the name and description from the code table 
                    TableValues = DataCodeValues[member] 
                    table_value = 0
                    if TableValues != 0: 
                        table_name = TableValues[0] # value type  
                        table_value = TableValues[value]  # text matching [value] code number	
                        if table_value == 0: 
                            table_value = "Undefined!" 
                        self.rawDataSheet.write(self.raw_sheet_row , dtt_col_number + 1, table_name)  # write text tag   
                        self.rawDataSheet.write(self.raw_sheet_row , dtt_col_number + 2, table_value)  # write text matching [value] code number			
			
                self.raw_sheet_row = self.raw_sheet_row + 1 

			
                # copy to function local variables for use in chart sheet below 
                if member == "dut_config": 
                    dut_config1 = value 
                    dut_config_text1 = text 
                elif  member == "dut_productSn": 
                    dut_sn1 = value 
                    dut_sn_text1 = text 
                elif  member == "dut_tag": 
                    dut_tag1 = value 
                    dut_tag_text1 = text 
                elif  member == "dut_port": 
                    dut_port1 = value 
                    dut_port_text1 = text 
                elif  member == "date_time": 
                    date_time1 = value  
                    date_time_text1 = text 
                elif  member == "load_on_seconds": 
                    load_on_seconds1 = value 
                    load_on_seconds_text1 = text 
                elif  member == "load_pre_idle_seconds": 
                    load_pre_idle_seconds1 = value 
                    load_pre_idle_seconds_text1 = text 
                elif  member == "set_loop": 
                    set_loop1 = value 
                    set_loop_text1 = text 
                elif  member == "test_step": 
                    test_step1 = value 
                    test_step_text1 = text 
                elif  member == "load_frame": 
                    load_frame1 = value 
                    load_frame_text1 = text 
                elif  member == "load_model": 
                    load_model1 = value 
                    load_model_text1 = text 
                elif  member == "load_channel": 
                    load_channel1 = value 
                    load_channel_text1 = text 
                elif  member == "test_code": 
                    test_code1 = value 
                    test_code_text1 = text 
                    test_type1 = table_value 
                    test_type_text1 = table_name 
                elif  member == "sample_count": 
                    sample_count1 = value 
                    sample_count_text1 = text 
                elif  member == "no_load_volt_post": 
                    no_load_volt_post1 = loadFunctions_Chroma_63600.roundNumber(self, value, 1e-3) # round to nearest mV
                    no_load_volt_post_text1 = text 
                elif  member == "test_group": 
                    test_group1 = value 
                    test_group_text1 = text 
        #End data table tags loop

        # header column 
        self.rawDataSheet.write(self.raw_sheet_row, value_name_col_number, "Time []")
        raw_sheet_time_row = self.raw_sheet_row
        self.rawDataSheet.write(self.raw_sheet_row + 1, value_name_col_number, "Voltage []")
        raw_sheet_volts_row = self.raw_sheet_row + 1
        self.rawDataSheet.write(self.raw_sheet_row+2, value_name_col_number, "Current []")
        raw_sheet_current_row = self.raw_sheet_row + 2
        # Data in rows 
        #LUA code used .Range for this functionality here we use .write and loop
        keyLoop = 1
        for key in T_Data['Time'].keys():
            self.rawDataSheet.write(self.raw_sheet_row, value_name_col_number + keyLoop, T_Data['Time'][key])
            keyLoop = keyLoop + 1
        keyLoop = 1
        for key in T_Data['Volts'].keys():
            self.rawDataSheet.write(self.raw_sheet_row + 1, value_name_col_number + keyLoop, T_Data['Volts'][key])
            keyLoop = keyLoop + 1
        keyLoop = 1
        for key in T_Data['Current'].keys():
            self.rawDataSheet.write(self.raw_sheet_row + 2, value_name_col_number + keyLoop, T_Data['Current'][key])
            keyLoop = keyLoop + 1
        self.rawDataSheet.set_column(0, 0, 24)
        self.rawDataSheet.set_column(1, 3, 17)
        self.raw_sheet_row = self.raw_sheet_row + 4 # leave a blank row 
        
        if new_chart_sheet:
            if new_excel:
                self.chart_sheet_number = 1
            else:
                self.chart_sheet_number = self.chart_sheet_number + 1 
            self.str_chart_sheet_name = str_chart_sheet_name_prefix + str(self.chart_sheet_number)
            active_Chart_Sheet = self.workBook.add_worksheet(self.str_chart_sheet_name)
            active_Chart_Sheet.set_column(1,1,32)
            active_Chart_Sheet.set_column(2,2,55)
            self.chart_sheet_row = 1 
            self.chart_sheet_col = 1 
        else:
            active_Chart_Sheet = self.workBook.get_worksheet_by_name(self.str_chart_sheet_name)
        
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, test_type_text1 + ":", self.headerCellFormat)
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, test_type1, self.headerCellFormat)

        if test_group1 != 0:
            active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 6, test_group_text1 + ":", self.headerCellFormat)
            active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 7, test_group1, self.headerCellFormat)
        
        chart_row = self.chart_sheet_row # add chart just under top header rows (above) 
        chart_col = self.chart_sheet_col + 3  # add chart just to right of statistics (below) 
        
        timeRange = xl_range(raw_sheet_time_row, 1, raw_sheet_time_row, self.max_samples)
        print "time range " + timeRange
        voltRange = xl_range(raw_sheet_volts_row, 1, raw_sheet_volts_row, self.max_samples)
        print "volt range " + voltRange
        currentRange = xl_range(raw_sheet_current_row, 1, raw_sheet_current_row, self.max_samples)
        print "current range " + currentRange

        # Create a new chart object.
        chart = self.workBook.add_chart({'type': 'scatter'})
        
        # Add a series to the chart.
        chart.add_series({'name': 'Voltage','categories': '=' + str_raw_sheet_name + '!' +  timeRange, 
                          'values': '=' + str_raw_sheet_name + '!' +  voltRange, 'marker': {'type': 'none'}, 'y2_axis': 1, 'line': {'color': color_volt, 'width': 1.25}})
        chart.add_series({'name': 'Current','categories': '=' + str_raw_sheet_name + '!' +  timeRange,
                          'values': '=' + str_raw_sheet_name + '!' +  currentRange, 'marker': {'type': 'none'}, 'line': {'color': color_curr, 'width': 1.25}})
# 
#        # Add a chart title and some axis labels.
        chart.set_x_axis({'name': 'Seconds', })
        chart.set_y_axis({'name': 'Current (A)', 'major_gridlines': {'visible': 0}, 'name_font': {'color': color_curr}})
        chart.set_y2_axis({'name': 'Voltage (V)', 'name_font': {'color': color_volt}})       
        chart.set_x_axis({'crossing': -4})        # Insert the chart into the worksheet.
        chart.set_size({'width': 800, 'height': 288})
        active_Chart_Sheet.insert_chart(chart_row, chart_col, chart)

        self.chart_sheet_row = self.chart_sheet_row + 1

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, dut_port_text1 + ":", self.headerCellFormat)
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, dut_port1, self.headerCellFormat)
        self.chart_sheet_row = self.chart_sheet_row + 1

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 4, dut_config_text1 + ":")
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 5, dut_config1,)
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 6, dut_sn_text1 + ":")
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 7, dut_sn1)
        self.chart_sheet_row = self.chart_sheet_row + 1

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, date_time_text1 + ":")
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, date_time1)
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 4, load_channel_text1 + ":")
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 5, load_channel1)
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 6, load_model_text1 + ":")
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 7, load_model1)
        self.chart_sheet_row = self.chart_sheet_row + 1

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 6, load_frame_text1 + ":")
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 7, load_frame1)
        self.chart_sheet_row = self.chart_sheet_row + 1

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col , "Pre-test no-load voltage (V): " , self.headerCellFormat)
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['no_load_volt_pre'], self.headerCellFormat)
        self.chart_sheet_row = self.chart_sheet_row + 1

#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, no_load_volt_post_text1 + ":", self.headerCellFormat)
#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, no_load_volt_post1, self.headerCellFormat)
#        self.chart_sheet_row = self.chart_sheet_row + 2

#	if math.abs(no_load_volt_post1 / T_Stats.no_load_volt_pre) < compare_match_factor then 
#		this_cell.Interior.Color = 0xB0B0FF -- light red (BBGGRR) 
#	end 

	
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "Constant current setting (A): ")
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Data['load_set_cc'])
        self.chart_sheet_row = self.chart_sheet_row + 1

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "CC average current (A): " , self.headerCellFormat)
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['cc_curr_ave'], self.headerCellFormat)
        self.chart_sheet_row = self.chart_sheet_row + 1

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "CC average voltage (V): "  , self.headerCellFormat)
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['cc_volt_ave'], self.headerCellFormat)
        self.chart_sheet_row = self.chart_sheet_row + 1

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "CC duration (sec): "  , self.headerCellFormat)
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['cc_time'], self.headerCellFormat)
        self.chart_sheet_row = self.chart_sheet_row + 2

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "Peak current, 1 sample (A): " )
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['peak_curr'])
        self.chart_sheet_row = self.chart_sheet_row + 1

        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "Max current (A): ")
        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['max_curr'])
        self.chart_sheet_row = self.chart_sheet_row + 3

#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "Trip current spec (A): ")
#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['trip_curr_spec'])
#        self.chart_sheet_row = self.chart_sheet_row + 1
#
#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "Time to trip (sec): "  , self.headerCellFormat)
#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['trip_time'], self.headerCellFormat)
#        self.chart_sheet_row = self.chart_sheet_row + 2
#
#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "Final current (A): "  , self.headerCellFormat)
#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['final_curr'], self.headerCellFormat)
#        self.chart_sheet_row = self.chart_sheet_row + 1
#
#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col, "Final voltage (V): ")
#        active_Chart_Sheet.write(self.chart_sheet_row, self.chart_sheet_col + 1, T_Stats['final_volt'])
#        self.chart_sheet_row = self.chart_sheet_row + 2

#	for loop = 0, 7, 2 do -- fit & align left/right column pairs 
#		sheet_chart.Columns(sheet_chart_col + loop):AutoFit() 
#		sheet_chart.Columns(sheet_chart_col + loop).HorizontalAlignment = -4152 -- right 
#		sheet_chart.Columns(sheet_chart_col + loop +1):AutoFit()
#		sheet_chart.Columns(sheet_chart_col + loop +1).HorizontalAlignment = -4131 -- left 
#	end -- for loop 
#	sheet_chart.Columns(sheet_chart_col + 1).ColumnWidth = 17 -- reset width that is too big from above AutoFit() (TEMP, replace with better autofit?) 
        

    def calcOCPData(self, T_Data, port):
        T_Stats = {}
        time_margin = 200e-6  # see also cc_slew_rate (NOTE: The time_margin should be greater than the rise/fall time.) 
        sample_margin = int(time_margin / self.sample_Period) # number of samples to exclude from boundary of averaging sample range (for calc high, low levels) 
        margin_percent = 3.0 # constant current samples allow range +/- percent  -- 2015-08-07 changed; was 2.0 
        short_cc_min = 15
		
        if port.find("USB") == -1: 
            tripped_current_max = 2.0   
            tripped_voltage_max = 0.3 
        elif port.find("AUX 5V") == -1:
            tripped_current_max =  0.5 
            tripped_voltage_max =  0.3 
        elif port.find("AUX 12V") == -1:
            tripped_current_max =  0.5 
            tripped_voltage_max =  0.3 
        elif port.find("HDMI 5V") == -1:
            tripped_current_max = 0.5 
            tripped_voltage_max = 0.3 
        elif port.find("IR 5V") == -1:
            tripped_current_max =  0.7 
            tripped_voltage_max =  0.3 
        else: 
            print(" *** NOTICE - Port not recognized. No over-current protection limits defined! ***") 
            tripped_current_max = 0 
            tripped_voltage_max = 0 
	
        # find time zero index (when load switched on) 
        time_zero_index = None
        for loop in range(0, self.max_samples): 
            if T_Data['Time'][loop] == 0: 
                time_zero_index = loop
                print "Time Zero Index " + str(time_zero_index)
                print "Trigger Time " + str(time_zero_index*self.sample_Period)
                break 
	
        if time_zero_index == None: 
            print("ERROR - Time zero not found!") 

        # find average high voltage (no-load)
        high_volt_ave = 0
        for loop in range(1, time_zero_index - sample_margin): 
            high_volt_ave = high_volt_ave + T_Data['Volts'][loop] 
            high_volt_n = loop
	
        high_volt_ave = high_volt_ave / high_volt_n  
        T_Stats['no_load_volt_pre'] = loadFunctions_Chroma_63600.roundNumber(self, high_volt_ave, 0.001) # round to nearest mV   

	
        # find peak current (Note: This works over entire sample.) 
        curr_peak = 0 
        curr_peak_index = 0
        for loop in range (1, self.max_samples): 
            if T_Data['Current'][loop] > curr_peak: 
                curr_peak = T_Data['Current'][loop] 
                curr_peak_index = loop 

        T_Stats['peak_curr'] = loadFunctions_Chroma_63600.roundNumber(self, curr_peak, 0.001) # round to nearest mA 

        # Find median current in median_sample_size samples starting at curr_peak_index.  
        # This rejects typical overshoot or small capacitance sourced peaks. 
        median_sample_size = 5 # odd values ideal 
        median_curr_peak = 0
        T_Sort = {} 
        sort_size = 0
        # transfer samples to a sort table 
        for loop in range(curr_peak_index, curr_peak_index + median_sample_size - 1): 
            if loop >= self.max_samples: 
                break # don't go past the end of source data 
            sort_size = sort_size + 1 
            T_Sort[sort_size] = T_Data['Current'][loop] 

        # sort T_Sort[] from low to high 
        k = int((sort_size + 1) / 2 ) # "+ 1" favors higher value for even sample sizes 
        for loop in range(1, k): 
            min_index = loop 
            min_value = T_Sort[loop] 
            for loop2 in range(loop + 1, sort_size): 
                if T_Sort[loop2] < min_value: 
                    min_index = loop2 
                    min_value = T_Sort[loop2] 
		
        T_Sort[loop], T_Sort[min_index] = T_Sort[min_index], T_Sort[loop] # swap values 
        median_curr_peak = T_Sort[k] # (Note: index in original data not tracked. Not sure if important.) 
        T_Stats['max_curr'] = loadFunctions_Chroma_63600.roundNumber(self, median_curr_peak, 0.001) # round to nearest mA 


	
        # find constant current interval 
        cc_index_1 = time_zero_index + sample_margin 
        cc_index_2 = cc_index_1 
        cc_curr_ave = 0 
        cc_volt_ave = 0 # average voltage during constant current state 
        cc_n = 0 
        cc_time = 0 
        cc_load = T_Data['load_set_cc'] 
        
	
        if not cc_load: 
            cc_load = loadFunctions_Chroma_63600.roundNumber(self, median_curr_peak, 0.1) 
            print("The data file doesn't contain the constant current load setting, so using max current rounded to nearest 0.1 amps = " + str(cc_load))

        if cc_load == "short": 
            cc_min = short_cc_min  
            cc_max = 100 
        else: 
            cc_min = cc_load * (1 - margin_percent / 100)
            cc_max = cc_load * (1 + margin_percent / 100) 

	
        for loop in range (cc_index_1, self.max_samples):
            this_curr = T_Data['Current'][loop] 
            this_volt = T_Data['Volts'][loop] 
            if this_curr < cc_min or this_curr > cc_max:
                break 
            else: 
                cc_curr_ave = cc_curr_ave + this_curr 
                cc_volt_ave = cc_volt_ave + this_volt 
                cc_index_2 = loop 

        if cc_index_2 == cc_index_1: 
            print("While analyzing resuls - Test specified constant current level not found.") 
        else: 
            cc_n = cc_index_2 - cc_index_1 + 1 
            cc_curr_ave = cc_curr_ave / cc_n 
            cc_volt_ave = cc_volt_ave / cc_n 
            cc_time = T_Data['Time'][cc_index_2] # assuming load starts at time = 0   

        T_Stats['cc_time'] =  loadFunctions_Chroma_63600.roundNumber(self, cc_time, 0.0001)  # rounding value should be variable or not rounded?  
        T_Stats['cc_curr_ave'] = loadFunctions_Chroma_63600.roundNumber(self, cc_curr_ave, 0.001) 
        T_Stats['cc_volt_ave'] = loadFunctions_Chroma_63600.roundNumber(self, cc_volt_ave, 0.001) 


	
        # find time to tripped_current_max 
        trip_index = 0 
        trip_time = 0
        for loop in range(time_zero_index + sample_margin, self.max_samples): 
            this_curr = T_Data['Current'][loop] 
            if this_curr >= tripped_current_max: # finds point where current stays below tripped_current_max 
                trip_index = 0 
            elif trip_index == 0: 
                this_volt = T_Data['Volts'][loop] 
                if this_volt <= tripped_voltage_max: 
                    trip_index = loop 

        if trip_index == 0:
            print("Specified tripped current level not found.") 
        else: 
            trip_time = T_Data['Time'][trip_index] # assuming load starts at time = 0 
	
        T_Stats['tripped_curr_spec'] = tripped_current_max 
        if trip_time != 0: 
            T_Stats['trip_time'] = loadFunctions_Chroma_63600.roundNumber(self, trip_time, 0.0001) # rounding value should be variable or not rounded? 
        else: 
            T_Stats['trip_time'] = "none" 

        # find final current (last x samples) 
        final_n = 10 
        final_curr_ave = 0 
        final_volt_ave = 0 
        for loop in range(self.max_samples - final_n + 1, self.max_samples):
            final_curr_ave = final_curr_ave + T_Data['Current'][loop] 
            final_volt_ave = final_volt_ave + T_Data['Volts'][loop] 
	
        final_curr_ave = final_curr_ave / final_n 
        final_volt_ave = final_volt_ave / final_n 
        T_Stats['final_curr'] = loadFunctions_Chroma_63600.roundNumber(self, final_curr_ave, 0.001) 
        T_Stats['final_volt'] = loadFunctions_Chroma_63600.roundNumber(self, final_volt_ave, 0.001)  
	
        # add calculation date time stamp, if this was read from a raw data file after test was run.
#        if T_Data['file_name'] != None: 
#            str_date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#            T_Stats['date_time'] = str_date_time 
	
        return T_Stats 


       










        
        
        
        
        
        
