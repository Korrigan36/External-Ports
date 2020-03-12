# -*- coding: utf-8 -*-
"""
Created on Wed Nov 08 12:32:35 2017

@author: v-stpurc
"""
import visa
import struct
import chromaLib
from time import sleep
from datetime import datetime

def constantCurrentTestMultPulse(self, load_frame_id, T_TestGroup):
        
    current_range_margin_factor = 0.9
    base_precision_volt = 1e-5 # precision for raw data
    base_precision_curr = 1e-5 # precision for raw data
    base_precision_time = 1e-6

    samples = self.max_samples
        
    last_load_channel = 10

    str_date_time_on = None
    str_date_time_off = None
    os_time_on = None
    os_time_off = None
    load_on_seconds = None
    
    return_status = False 
    early_turnon_count = 0 # default
    early_turnon_delay = 2000 # ms, constant 
    EarlyTurnonList = {}
    T_ResultsGroup = {}
    
    test_count = len(T_TestGroup )
    test_group = T_TestGroup[0]
    
    for loop in range (1, last_load_channel):
        chromaLib.load_SetActiveOff(self.load, str(loop))
   
    #Set up the Chroma loads for a multiple simultaneous load test
    #go through each 'test'. In this case each test is a different load on it's own port
    for loop in range (1, test_count):
        T_ResultsGroup[loop] = newType_DataTable(self)
        T_Results = T_ResultsGroup[loop] 
        T_Test = T_TestGroup[loop] 
        test_id = T_Test['TestID']
        T_Results['test_code'] = test_id
        test_port = T_Test['Port']
        load_num = T_Test['Channel']
        #current is used  for the dynamic load L2, pulse current is dynamic load L1
        current = T_Test['Amps']
        duration = T_Test['Duration']
        trigger_time = T_Test['Trigger']
        current_rise_slew_rate = T_Test['RiseSlewRate']
        current_fall_slew_rate = T_Test['FallSlewRate']
        pulse_t1 = T_Test['PulseT1']
        pulse_t1_current = T_Test['PulseT1Amps']
        pulse_t2 = T_Test['PulseT2']
        pulse_count = T_Test['PulseCount']
        if pulse_count == None:
            pulse_count = 1
        if pulse_t1 != None:
            pulse_mode_enable = True
        else:
            pulse_mode_enable = False
            
        dig_trig_source = "LOADON"
        console = T_Test['Console']
#            dut_tag = DutInfo[console]['dut_tag'] 
        sample_period = float(duration) / float(samples)
        sample_period = roundNumber(self, sample_period, self.sample_period_resolution)
		
        if sample_period > self.max_sample_period:
            print "Sample period too long, setting to max: " 
            sample_period = self.max_sample_period 
		
        duration = samples * sample_period # calculate actual duration 
        
        trigger_sample = int(trigger_time / sample_period + 1) # trigger_sample range 1~4096 (or samples?)
	
        if trigger_sample > samples: 
            trigger_sample = samples 
            
#        chromaLib.load_TurnOffParallelMode(self.load)
        chromaLib.load_SelectSingleChannel(self.load, load_num)
        chromaLib.load_EnableSingleChannel(self.load, "ON")
        chromaLib.load_TurnOnOffLoad(self.load, "OFF")
        chromaLib.load_TurnOnOffLoadShort(self.load, "OFF")
        chromaLib.load_AllRunOnOff(self.load, "OFF")
        chromaLib.load_AutoOnOff(self.load, "OFF")
        chromaLib.load_ConfigWindow(self.load, 0.10)
        chromaLib.load_ConfigParallel(self.load, "NONE")
        chromaLib.load_ConfigSync(self.load, "NONE")
        chromaLib.load_VoltLatchOnOff(self.load, "OFF")
        chromaLib.load_VoltOff(self.load, 0.0)
        chromaLib.load_VoltOn(self.load, 0.0)
        chromaLib.load_VoltSign(self.load, "PLUS")
            
        str_module_id = self.load.query("CHAN:ID?")
        str_module_id = str_module_id.replace (',','_')
        str_module_id = str_module_id.replace ('\n','')
        str_module_id = str_module_id.replace ('\r','')
        
        # select left/right display for dual channel modules (NOTE: If both L&R used, then the last one set up is shown.) 
        if str_module_id.find("L_") != -1: 
            self.load.write("SHOW:DISP L") 
        elif str_module_id.find("R_") != -1:
            self.load.write("SHOW:DISP R") 
            
				
        LoadSpec = getLoadSpecs(self, str_module_id)
        max_cch_amps        = LoadSpec['cch_amps_max'] 
        cch_slew_max_rate   = LoadSpec['cch_slew_max'] 
        cch_slew_min_rate   = LoadSpec['cch_slew_min'] 
        max_ccm_amps        = LoadSpec['ccm_amps_max'] 
        ccm_slew_max_rate   = LoadSpec['ccm_slew_max'] 
        ccm_slew_min_rate   = LoadSpec['ccm_slew_min']
        max_ccl_amps        = LoadSpec['ccl_amps_max'] 
        ccl_slew_max_rate   = LoadSpec['ccl_slew_max']
        ccl_slew_min_rate   = LoadSpec['ccl_slew_min'] 
        
        if current > (max_ccm_amps * current_range_margin_factor):
            if pulse_mode_enable:
                str_current_range = "CCDH"
            else:
                str_current_range = "CCH"
            #Slew at max rate, Chroma is slow enough. Change if tests require
            cc_rise_slew_rate = cch_slew_max_rate
            cc_fall_slew_rate = cch_slew_max_rate
        elif current > (max_ccl_amps * current_range_margin_factor):
            if pulse_mode_enable:
                str_current_range = "CCDM"
            else:
                str_current_range = "CCM"
            cc_rise_slew_rate = ccm_slew_max_rate
            cc_fall_slew_rate = ccm_slew_max_rate
        else:
            if pulse_mode_enable:
                str_current_range = "CCDL"
            else:
                str_current_range = "CCL"
            cc_rise_slew_rate = ccl_slew_max_rate
            cc_fall_slew_rate = ccl_slew_max_rate
                
        
        #Check Port IV range was here
        #Check Slew Rate was here
                
        print("Sampling rate = " + str(1 / sample_period) + " Hz for " + str(duration) + " seconds.") 

        chromaLib.load_SelectSingleChannel(self.load, load_num)
        chromaLib.load_CurrentMode(self.load, str_current_range)
        chromaLib.load_VoltRange(self.load, "LOW")
        
        #pulse_mode_enable is determined by a value in Pulse T1 which is the dynamic load off time.
        #T1 current is low typically 0.15A
        #The actual dynamic load is T2 and T2 is on for a short time. That load is given in the Amps column, typically 4A
        if pulse_mode_enable:
            T_Results['top_axis'] = "current"
            chromaLib.load_DynamicRiseTime(self.load, str(cc_rise_slew_rate))
            chromaLib.load_DynamicFallTime(self.load, str(cc_fall_slew_rate))
            chromaLib.load_DynamicLoad1(self.load, str(pulse_t1_current))
            chromaLib.load_DynamicTime1(self.load, str(pulse_t1))
            chromaLib.load_DynamicLoad2(self.load, str(current))
            chromaLib.load_DynamicTime2(self.load, str(pulse_t2))
            chromaLib.load_DynamicPulseCount(self.load, str(pulse_count))
#            chromaLib.load_ConfigParallel(self.load, "MASTER")
        else:
            chromaLib.load_RiseTime(self.load, str(cc_rise_slew_rate))
            chromaLib.load_FallTime(self.load, str(cc_fall_slew_rate))
            chromaLib.load_SetCurrent(self.load, current)
#            chromaLib.load_ConfigParallel(self.load, "SLAVE")
            
            #Special case pulse_t1 = 0, pre digitizing turn on of load was here
            
        # Set up Digitizing mode 
        chromaLib.load_DigSampleTime(self.load, sample_period)  # [2us~40ms, resolution 2us] set sampling time  
        chromaLib.load_DigSamplePoints(self.load, samples) # [1~4096, default 4096] set sampling points 
        #We set the trigger sample at 2ms so there is 2ms of data preceeding the load turn on.
        chromaLib.load_DigTriggerPoint(self.load, trigger_sample) # [1~4096, default 2000] set trigger point
        
        
        rm = visa.ResourceManager()
        lib = rm.visalib

        chromaLib.load_DigTriggerSource(self.load, "LOADON") # [default LOADON] set trigger source
#        chromaLib.load_DigTriggerSource(self.load, "BUS") # [default LOADON] set trigger source
        self.load.timeout = 2500
 		
        T_Results['load_set_cc'] = current 
        T_Results['sample_period'] = sample_period 
        T_Results['load_channel'] = load_num 
        T_Results['load_model'] = str_module_id  
        T_Results['load_frame'] = load_frame_id
        T_Results['dut_port'] = test_port  
        T_Results['trigger_sample'] = trigger_sample 
        
        #End of for loop 
        #We've set up each load in the group now. Typically there are three loads in a group
        #When we say RUN below all loads turn on. The one set for dynamic load does it's thing instead of CC mode.
        #The dynamic load is off for T1 2ms then it's on for T2 10us
        
    AllDataV = {}
    AllDataI = {}
    
    #Early turn on delay was here
    
    print "Starting Test"
    for loop in range (1, test_count):
        T_Test = T_TestGroup[loop]
        load_num = T_Test['Channel']
        chromaLib.load_SelectSingleChannel(self.load, load_num)
        # Acquire V/I waveforms at load turn-on 
        self.load.write("DIG:INIT") # initialize digitizing 

    #Wait for all loads to be in wait for trigger mode
    while True:
        for loop in range (1, test_count):
            load_num = T_TestGroup[loop]['Channel']
            chromaLib.load_SelectSingleChannel(self.load, load_num)
            returnString = self.load.query("DIG:TRIG?") 
            if returnString.find('WAIT_TRIG') != -1:
                is_ready = True
            else:
                is_ready = False
                
        if is_ready == True:
            break
        
    os_time_on = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("Test started time: " + os_time_on) 
    
    #Early turn on stuff was here
    
    #For multiple loads RUN turns them all on
    self.load.write("RUN")

    #Wait for all loads to be in idle trigger mode
    while True:
        for loop in range (1, test_count):
            load_num = T_TestGroup[loop]['Channel']
            chromaLib.load_SelectSingleChannel(self.load, load_num)
            returnString = self.load.query("DIG:TRIG?") 
            if returnString.find('IDLE') != -1:
                is_ready = True
            else:
                is_ready = False
                
        if is_ready == True:
            break

    #For multiple loads RUN turns them all on
    self.load.write("ABORT")
   
    os_time_off = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("Test finished time: " + os_time_off)
    
    #Confirm loads are off was here
    
    for loop in range (1, test_count):
        T_Results = T_ResultsGroup[loop]  
        load_num = T_Results['load_channel']
        chromaLib.load_SelectSingleChannel(self.load, load_num)
        meas_ave_samples = 10
        after_volt = 0
        for loop5 in range (1, meas_ave_samples):
            after_volt = after_volt + float(self.load.query("MEAS:VOLT?"))     
        after_volt = after_volt / meas_ave_samples # calculate average of samples
        T_Results['no_load_volt_post'] = roundNumber(self, after_volt, base_precision_volt)
        
    for loop in range (1, test_count):
        DataV = {} 
        DataI = {} 
        T_Results = T_ResultsGroup[loop] 
        load_num = T_Results['load_channel'] 
        chromaLib.load_SelectSingleChannel(self.load, load_num)
        # Read out waveforms  
        print("Reading Out Data.") 
				
        self.load.timeout = 30000
        while True:
            returnValue = self.load.query("DIG:WAV:CAP?")
            print(returnValue)
            if returnValue.find("OK") != -1:
                break
            else:
                sleep(5)
         
        DataV = self.load.query_binary_values("DIG:WAV:DATA? V", datatype='d', is_big_endian=False) # Largest buffer currently supported is 262144.
        DataI = self.load.query_binary_values("DIG:WAV:DATA? I", datatype='d', is_big_endian=False) # Largest buffer currently supported is 262144.
        
        AllDataV[loop] = DataV 
        AllDataI[loop] = DataI 
        
    for loop in range (1, test_count):
        DataV = AllDataV[loop]
        DataI = AllDataI[loop] 
        T_Results = T_ResultsGroup[loop]  

        T_Results['date_time'] = os_time_on
        T_Results['load_on_seconds'] = load_on_seconds
        T_Results['Volts'] = {}
        T_Results['Current'] = {}
        
        #Start time negative. Zero is trigger point
        time_marker = -sample_period * (trigger_sample - 1)
        
        for byteloop in range (0, samples/2):
            ba = bytearray(struct.pack("d", DataV[byteloop]))   
            data = [ba[3], ba[2], ba[1], ba[0]]
            b = ''.join(chr(i) for i in data)
            lowFloat = struct.unpack('>f', b)[0]
            T_Results['Volts'][byteloop*2]= roundNumber(self, lowFloat, base_precision_volt)
            data = [ba[7], ba[6], ba[5], ba[4]]
            b = ''.join(chr(i) for i in data)
            hiFloat = struct.unpack('>f', b)[0]
            T_Results['Volts'][byteloop*2 + 1]= roundNumber(self, hiFloat, base_precision_volt)

            ba = bytearray(struct.pack("d", DataI[byteloop]))   
            data = [ba[3], ba[2], ba[1], ba[0]]
            b = ''.join(chr(i) for i in data)
            lowFloat = struct.unpack('>f', b)[0]
            T_Results['Current'][byteloop*2]= roundNumber(self, lowFloat, base_precision_curr)
            data = [ba[7], ba[6], ba[5], ba[4]]
            b = ''.join(chr(i) for i in data)
            hiFloat = struct.unpack('>f', b)[0]
            T_Results['Current'][byteloop*2 + 1]= roundNumber(self, hiFloat, base_precision_curr)
            
        #Increment time_marker
        for byteloop in range (0, samples):
            T_Results['Time'][byteloop] = roundNumber(self, time_marker, base_precision_time)
            time_marker = time_marker + sample_period
            
        T_Results['sample_count'] = samples 
        T_Results['test_gropu'] = test_group
       
    return T_ResultsGroup, sample_period

def thresholdTest(self, load_frame_id, T_TestGroup):
        
    current_range_margin_factor = 0.9
    base_precision_volt = 1e-5 # precision for raw data
    base_precision_curr = 1e-5 # precision for raw data
    base_precision_time = 1e-6

    samples = self.max_samples
        
    last_load_channel = 10

    str_date_time_on = None
    str_date_time_off = None
    os_time_on = None
    os_time_off = None
    load_on_seconds = None
    
    return_status = False 
    early_turnon_count = 0 # default
    early_turnon_delay = 2000 # ms, constant 
    EarlyTurnonList = {}
    T_ResultsGroup = {}
    
    test_count = len(T_TestGroup )
    test_group = T_TestGroup[0]
    
    #The following must work because max is an integer and Amps is converted to integer thus anything less than one looks like zero?
    max = 0
    ocp_test = -1
    #This seems to limit the number grouped to three
#    for index in range(1,4):
    for index in range(1,test_count):
        print ("got here*****" + str(index))
        print (T_TestGroup[index])
        if long(T_TestGroup[index]['Amps']) > max:
            ocp_test = index
            print ("OCP Test: " + str(ocp_test))
    print("OCP test on " + T_TestGroup[ocp_test]["Port"])
    
    for loop in range (1, last_load_channel):
        chromaLib.load_SetActiveOff(self.load, str(loop))
   
    #Set up the Chroma loads for a multiple simultaneous load test
    #go through each 'test'. In this case each test is a different load on it's own port
    for loop in range (1, test_count):
        T_ResultsGroup[loop] = newType_DataTable(self)
        T_Results = T_ResultsGroup[loop] 
        T_Test = T_TestGroup[loop] 
        test_id = T_Test['TestID']
        T_Results['test_code'] = test_id
        test_port = T_Test['Port']
        load_num = T_Test['Channel']
        #current is used  for the dynamic load L2, pulse current is dynamic load L1
        current = T_Test['Amps']
        duration = T_Test['Duration']
        trigger_time = T_Test['Trigger']
            
        dig_trig_source = "LOADON"
        console = T_Test['Console']
#            dut_tag = DutInfo[console]['dut_tag'] 
        sample_period = float(duration) / float(samples)
        sample_period = roundNumber(self, sample_period, self.sample_period_resolution)
		
        if sample_period > self.max_sample_period:
            print "Sample period too long, setting to max: " 
            sample_period = self.max_sample_period 
		
        duration = samples * sample_period # calculate actual duration 
        
        trigger_sample = int(trigger_time / sample_period + 1) # trigger_sample range 1~4096 (or samples?)
	
        if trigger_sample > samples: 
            trigger_sample = samples 
            
#        chromaLib.load_TurnOffParallelMode(self.load)
        chromaLib.load_SelectSingleChannel(self.load, load_num)
        chromaLib.load_EnableSingleChannel(self.load, "ON")
        chromaLib.load_TurnOnOffLoad(self.load, "OFF")
        chromaLib.load_TurnOnOffLoadShort(self.load, "OFF")
        chromaLib.load_AllRunOnOff(self.load, "OFF")
        chromaLib.load_AutoOnOff(self.load, "OFF")
        chromaLib.load_ConfigWindow(self.load, 0.10)
        chromaLib.load_ConfigParallel(self.load, "NONE")
        chromaLib.load_ConfigSync(self.load, "NONE")
        chromaLib.load_VoltLatchOnOff(self.load, "OFF")
        chromaLib.load_VoltOff(self.load, 0.0)
        chromaLib.load_VoltOn(self.load, 0.0)
        chromaLib.load_VoltSign(self.load, "PLUS")
            
        str_module_id = self.load.query("CHAN:ID?")
        str_module_id = str_module_id.replace (',','_')
        str_module_id = str_module_id.replace ('\n','')
        str_module_id = str_module_id.replace ('\r','')
        
        # select left/right display for dual channel modules (NOTE: If both L&R used, then the last one set up is shown.) 
        if str_module_id.find("L_") != -1: 
            self.load.write("SHOW:DISP L") 
        elif str_module_id.find("R_") != -1:
            self.load.write("SHOW:DISP R") 
            
				
        LoadSpec = getLoadSpecs(self, str_module_id)
        max_cch_amps        = LoadSpec['cch_amps_max'] 
        cch_slew_max_rate   = LoadSpec['cch_slew_max'] 
        cch_slew_min_rate   = LoadSpec['cch_slew_min'] 
        max_ccm_amps        = LoadSpec['ccm_amps_max'] 
        ccm_slew_max_rate   = LoadSpec['ccm_slew_max'] 
        ccm_slew_min_rate   = LoadSpec['ccm_slew_min']
        max_ccl_amps        = LoadSpec['ccl_amps_max'] 
        ccl_slew_max_rate   = LoadSpec['ccl_slew_max']
        ccl_slew_min_rate   = LoadSpec['ccl_slew_min'] 
        
        if current > (max_ccm_amps * current_range_margin_factor):
            if loop == ocp_test:
                str_current_range = "OCPH"
            else:
                str_current_range = "CCH"
            #Slew at max rate, Chroma is slow enough. Change if tests require
            cc_rise_slew_rate = cch_slew_max_rate
            cc_fall_slew_rate = cch_slew_max_rate
        elif current > (max_ccl_amps * current_range_margin_factor):
            if loop == ocp_test:
                str_current_range = "OCPM"
            else:
                str_current_range = "CCM"
            cc_rise_slew_rate = ccm_slew_max_rate
            cc_fall_slew_rate = ccm_slew_max_rate
        else:
            if loop == ocp_test:
                str_current_range = "OCPL"
            else:
                str_current_range = "CCL"
            cc_rise_slew_rate = ccl_slew_max_rate
            cc_fall_slew_rate = ccl_slew_max_rate
                
        
        #Check Port IV range was here
        #Check Slew Rate was here
                
        print("Sampling rate = " + str(1 / sample_period) + " Hz for " + str(duration) + " seconds.") 

        chromaLib.load_SelectSingleChannel(self.load, load_num)
        chromaLib.load_CurrentMode(self.load, str_current_range)
        chromaLib.load_VoltRange(self.load, "LOW")
        
        if loop == ocp_test:
            T_Results['top_axis'] = "current"
            self.load.write("ADVANCE:OCP:ISTART 0")
            self.load.write("ADVANCE:OCP:IEND " + str(current))
            self.load.write("ADVANCE:OCP:STEP 1000")
            self.load.write("ADVANCE:OCP:DWELL 10ms")
            self.load.write("ADVANCE:OCP:TRIGGER:VOLTAGE 0.5")
        else:
            chromaLib.load_RiseTime(self.load, str(cc_rise_slew_rate))
            chromaLib.load_FallTime(self.load, str(cc_fall_slew_rate))
            chromaLib.load_SetCurrent(self.load, current)
#            chromaLib.load_ConfigParallel(self.load, "SLAVE")
            
            #Special case pulse_t1 = 0, pre digitizing turn on of load was here
            
        # Set up Digitizing mode 
        chromaLib.load_DigSampleTime(self.load, sample_period)  # [2us~40ms, resolution 2us] set sampling time  
        chromaLib.load_DigSamplePoints(self.load, samples) # [1~4096, default 4096] set sampling points 
        chromaLib.load_DigTriggerPoint(self.load, trigger_sample) # [1~4096, default 2000] set trigger point
        
        
        rm = visa.ResourceManager()
        lib = rm.visalib

        chromaLib.load_DigTriggerSource(self.load, "LOADON") # [default LOADON] set trigger source
#        chromaLib.load_DigTriggerSource(self.load, "BUS") # [default LOADON] set trigger source
        self.load.timeout = 2500
 		
        T_Results['load_set_cc'] = current 
        T_Results['sample_period'] = sample_period 
        T_Results['load_channel'] = load_num 
        T_Results['load_model'] = str_module_id  
        T_Results['load_frame'] = load_frame_id
        T_Results['dut_port'] = test_port  
        T_Results['trigger_sample'] = trigger_sample 
        
        #End of for loop 
    AllDataV = {}
    AllDataI = {}
    
    #Early turn on delay was here
    
    print "Starting Test"
    for loop in range (1, test_count):
        T_Test = T_TestGroup[loop]
        load_num = T_Test['Channel']
        chromaLib.load_SelectSingleChannel(self.load, load_num)
        # Acquire V/I waveforms at load turn-on 
        self.load.write("DIG:INIT") # initialize digitizing 

    #Wait for all loads to be in wait for trigger mode
    print("Waiting for DIG initialization")
    while True:
        for loop in range (1, test_count):
            load_num = T_TestGroup[loop]['Channel']
            chromaLib.load_SelectSingleChannel(self.load, load_num)
            returnString = self.load.query("DIG:TRIG?") 
            print(returnString)
            if returnString.find('WAIT_TRIG') != -1:
                is_ready = True
            else:
                is_ready = False
                
        if is_ready == True:
            break
        sleep(1)
        
    os_time_on = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("Test started time: " + os_time_on) 
    
    #Early turn on stuff was here
    
    #For multiple loads RUN turns them all on
    self.load.write("RUN")
#    returnValue = lib.assert_trigger(self.load.session, 0)
#    returnValue = self.load.assert_trigger()
#    self.load.wait_for_srq()
#    print "returnValue " + returnValue

    #Wait for all loads to be in idle trigger mode
    print("Wait for trigger")
    while True:
        for loop in range (1, test_count):
            load_num = T_TestGroup[loop]['Channel']
            chromaLib.load_SelectSingleChannel(self.load, load_num)
            returnString = self.load.query("DIG:TRIG?") 
            print(returnString)
            if returnString.find('IDLE') != -1:
                is_ready = True
            else:
                is_ready = False
                
        if is_ready == True:
            break
        sleep(2)

    #For multiple loads RUN turns them all on
    self.load.write("ABORT")
   
    os_time_off = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("Test finished time: " + os_time_off)
    
    #Confirm loads are off was here
    
    for loop in range (1, test_count):
        T_Results = T_ResultsGroup[loop]  
        load_num = T_Results['load_channel']
        chromaLib.load_SelectSingleChannel(self.load, load_num)
        meas_ave_samples = 10
        after_volt = 0
        for loop5 in range (1, meas_ave_samples):
            after_volt = after_volt + float(self.load.query("MEAS:VOLT?"))     
        after_volt = after_volt / meas_ave_samples # calculate average of samples
        T_Results['no_load_volt_post'] = roundNumber(self, after_volt, base_precision_volt)
        
    for loop in range (1, test_count):
        DataV = {} 
        DataI = {} 
        T_Results = T_ResultsGroup[loop] 
        load_num = T_Results['load_channel'] 
        chromaLib.load_SelectSingleChannel(self.load, load_num)
        # Read out waveforms  
        print("Reading Out Data.") 
				
        self.load.timeout = 30000
        while True:
            returnValue = self.load.query("DIG:WAV:CAP?")
            print(returnValue)
            if returnValue.find("OK") != -1:
                break
            else:
                sleep(5)
         
        DataV = self.load.query_binary_values("DIG:WAV:DATA? V", datatype='d', is_big_endian=False) # Largest buffer currently supported is 262144.
        DataI = self.load.query_binary_values("DIG:WAV:DATA? I", datatype='d', is_big_endian=False) # Largest buffer currently supported is 262144.
        
        AllDataV[loop] = DataV 
        AllDataI[loop] = DataI 
        
    for loop in range (1, test_count):
        DataV = AllDataV[loop]
        DataI = AllDataI[loop] 
        T_Results = T_ResultsGroup[loop]  

        T_Results['date_time'] = os_time_on
        T_Results['load_on_seconds'] = load_on_seconds
        T_Results['Volts'] = {}
        T_Results['Current'] = {}
        #T_Results["Trip"] = self.load.read("ADVANCE:OCP:RESULT?")
        
        #Start time negative. Zero is trigger point
        time_marker = -sample_period * (trigger_sample - 1)
        
        for byteloop in range (0, samples/2):
            ba = bytearray(struct.pack("d", DataV[byteloop]))   
            data = [ba[3], ba[2], ba[1], ba[0]]
            b = ''.join(chr(i) for i in data)
            lowFloat = struct.unpack('>f', b)[0]
            T_Results['Volts'][byteloop*2]= roundNumber(self, lowFloat, base_precision_volt)
            data = [ba[7], ba[6], ba[5], ba[4]]
            b = ''.join(chr(i) for i in data)
            hiFloat = struct.unpack('>f', b)[0]
            T_Results['Volts'][byteloop*2 + 1]= roundNumber(self, hiFloat, base_precision_volt)

            ba = bytearray(struct.pack("d", DataI[byteloop]))   
            data = [ba[3], ba[2], ba[1], ba[0]]
            b = ''.join(chr(i) for i in data)
            lowFloat = struct.unpack('>f', b)[0]
            T_Results['Current'][byteloop*2]= roundNumber(self, lowFloat, base_precision_curr)
            data = [ba[7], ba[6], ba[5], ba[4]]
            b = ''.join(chr(i) for i in data)
            hiFloat = struct.unpack('>f', b)[0]
            T_Results['Current'][byteloop*2 + 1]= roundNumber(self, hiFloat, base_precision_curr)
            
        #Increment time_marker
        for byteloop in range (0, samples):
            T_Results['Time'][byteloop] = roundNumber(self, time_marker, base_precision_time)
            time_marker = time_marker + sample_period
            
        T_Results['sample_count'] = samples 
        T_Results['test_gropu'] = test_group
        
        
    print("\nReset Xbox!")
    sleep(60)
    
       
    return T_ResultsGroup, sample_period



def constantCurrentTest2(self, load_frame_id, test_port, load_num, current, duration, trigger_time, current_rise_slew_rate, current_fall_slew_rate):
    
    samples = self.max_samples
    cc_rise_slew_rate = None
    cc_fall_slew_rate = None

    str_date_time_on = None
    str_date_time_off = None
    os_time_on = None
    os_time_off = None
    load_on_seconds = None
    load_pre_idle_seconds = None
    str_module_id = None 
    return_status = False 
    short_mode = False 
    after_curr = None
    after_volt = None 

    base_precision_volt = 1e-5 # precision for raw data
    base_precision_curr = 1e-5 # precision for raw data
    base_precision_time = 1e-6

    #T_Results is returned and becomes an instance of All_Results back in timerThread
    T_Results = newType_DataTable(self)
    #Add test code key to T_Results and set value
    T_Results['test_code'] = 1110
    
    current_range_margin_factor = 0.9
    
    #Check to see if current is a string
    #Add load_set_cc key and set value
    if isinstance(current, basestring):
        if current.lower() == "short":
            short_mode = True 
            T_Results['load_set_cc'] = "short" 
    else: 
        T_Results['load_set_cc'] = current 
 
	
    #Sample period = duration / max samples
    sample_period = float(duration) / samples # [2us~40ms, resolution 2us]
    sample_period = roundNumber(self, sample_period, self.sample_period_resolution) 
    self.sample_Period = sample_period
	
    if sample_period > self.max_sample_period:
        print("Sample period too long, setting to max: " + str(self.max_sample_period) + " seconds.") 
        sample_period = self.max_sample_period 
	
    #Recalculate duration because it may not match, rounding errors etc
    duration = samples * sample_period # calculate actual duration 
    trigger_sample = int(trigger_time / sample_period + 1) # trigger_sample range 1~4096 (or samples?)
	
    if trigger_sample > samples: 
        trigger_sample = samples 
	
    DataV = {}
    DataI = {} 

					
    # Common setup for all ports 					
					
    chromaLib.load_SelectSingleChannel(self.load, load_num) # select one load module  
    chromaLib.load_EnableSingleChannel(self.load, "ON") # [default ON] ??? Enable load module 
    chromaLib.load_TurnOnOffLoad(self.load, "OFF") # turn off load 
    chromaLib.load_TurnOnOffLoadShort(self.load, "OFF") #turn off short-circuit state  
    chromaLib.load_AllRunOnOff(self.load, "OFF") # [default ON] Turn off all load run (linked on/off) 
    chromaLib.load_AutoOnOff(self.load, "OFF") # [default OFF]  
    chromaLib.load_ConfigWindow(self.load, 0.10) # [default 0.02s] -- 2015-08-14 changed, only affects panel display 
    chromaLib.load_ConfigParallel(self.load, "NONE") # [default NONE] 
    chromaLib.load_ConfigSync(self.load, "NONE") # [default NONE] [Why set this?]  
    chromaLib.load_VoltLatchOnOff(self.load, "OFF") # [default OFF]
    chromaLib.load_VoltOff(self.load, 0.0) # [default 0.00] 
    chromaLib.load_VoltOn(self.load, 0.0) # [default 0.00] 
    chromaLib.load_VoltSign(self.load, "PLUS") # [default PLUS] 
	
    #This direct call to the load object should be in chromLib. Not sure why it's here?				
    str_module_id = self.load.query("CHAN:ID?")
    str_module_id = str_module_id.replace (',','_')
    str_module_id = str_module_id.replace ('\n','')
    str_module_id = str_module_id.replace ('\r','')
    
    # select left/right display for dual channel modules (NOTE: If both L&R used, then the last one set up is shown.) 
    if str_module_id.find("L_") != -1: 
        self.load.write("SHOW:DISP L") 
    elif str_module_id.find("R_") != -1:
        self.load.write("SHOW:DISP R") 
        
				
    LoadSpec = getLoadSpecs(self, str_module_id)
    max_cch_amps        = LoadSpec['cch_amps_max'] 
    cch_slew_max_rate   = LoadSpec['cch_slew_max'] 
    cch_slew_min_rate   = LoadSpec['cch_slew_min'] 
    max_ccm_amps        = LoadSpec['ccm_amps_max'] 
    ccm_slew_max_rate   = LoadSpec['ccm_slew_max'] 
    ccm_slew_min_rate   = LoadSpec['ccm_slew_min']
    max_ccl_amps        = LoadSpec['ccl_amps_max'] 
    ccl_slew_max_rate   = LoadSpec['ccl_slew_max']
    ccl_slew_min_rate   = LoadSpec['ccl_slew_min'] 
    
    if current > (max_ccm_amps * current_range_margin_factor):
        str_current_range = "CCH"
    elif current > (max_ccl_amps * current_range_margin_factor):
        str_current_range = "CCM"
    else:
        str_current_range = "CCL"
        

    #Check Port IV range was here
    #Check Slew Rate was here
    #Check current range was here. For now use low current

    print("Sampling rate = " + str(1 / sample_period) + " Hz for " + str(duration) + " seconds.") 
    print("Constant-current load = " + str(current) + " Amps, triggered at " + str((trigger_sample-1) * sample_period) + " seconds.")
     
    chromaLib.load_CurrentMode(self.load, str_current_range) 
    chromaLib.load_VoltRange(self.load, "LOW") 
    #Support for pulse mode was here
    #Set Slew rates to max for now, not used in config spreadsheet right now
    chromaLib.load_RiseTime(self.load, "MAX") 
    chromaLib.load_FallTime(self.load, "MAX") 
    
    if not short_mode: 
        chromaLib.load_SetCurrent(self.load, current) # set CC amps 
				
    # Set up Digitizing mode 
    chromaLib.load_DigSampleTime(self.load, sample_period)  # [2us~40ms, resolution 2us] set sampling time  
    chromaLib.load_DigSamplePoints(self.load, samples) # [1~4096, default 4096] set sampling points 
    chromaLib.load_DigTriggerPoint(self.load, trigger_sample) # [1~4096, default 2000] set trigger point
    chromaLib.load_DigTriggerSource(self.load, "LOADON") # [default LOADON] set trigger source
    # Acquire V/I waveforms at load turn-on 
    self.load.write("DIG:INIT") # initialize digitizing 
 		
    while True:
        returnString = self.load.query("DIG:TRIG?") 
        if returnString.find('WAIT_TRIG') != -1:
            break
            
    os_time_on = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("Test started time: " + os_time_on) 

				
    if short_mode: 
        chromaLib.load_TurnOnOffLoadShort(self.load, "ON")
    else: 
        chromaLib.load_TurnOnOffLoad(self.load, "ON")  
    
    while True:
        returnString = self.load.query("DIG:TRIG?") 
        if returnString.find("IDLE") != -1:
            break

    if short_mode: 
        chromaLib.load_TurnOnOffLoadShort(self.load, "OFF")
    else: 
        chromaLib.load_TurnOnOffLoad(self.load, "OFF")  
    
    os_time_off = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("Test finished time: " + os_time_off)
		
    # Read out waveforms  
    print("Reading Out Data.") 
				
    self.load.timeout = 30000
    while True:
        returnValue = self.load.query("DIG:WAV:CAP?")
        sleep(3)
        print returnValue
        if returnValue.find("OK") != -1:
            break
     
    DataV = self.load.query_binary_values("DIG:WAV:DATA? V", datatype='d', is_big_endian=False) # Largest buffer currently supported is 262144.
    DataI = self.load.query_binary_values("DIG:WAV:DATA? I", datatype='d', is_big_endian=False) # Largest buffer currently supported is 262144.

    T_Results['date_time'] = os_time_on
    T_Results['load_on_seconds'] = load_on_seconds
    T_Results['Volts'] = {}
    T_Results['Current'] = {}
    
    #Start time axis negative. At trigger point time will be zero
    time_marker = -sample_period * (trigger_sample - 1)
    
    for byteloop in range (0, samples/2):
        ba = bytearray(struct.pack("d", DataV[byteloop]))   
        data = [ba[3], ba[2], ba[1], ba[0]]
        b = ''.join(chr(i) for i in data)
        lowFloat = struct.unpack('>f', b)[0]
        T_Results['Volts'][byteloop*2]= roundNumber(self, lowFloat, base_precision_volt)
        data = [ba[7], ba[6], ba[5], ba[4]]
        b = ''.join(chr(i) for i in data)
        hiFloat = struct.unpack('>f', b)[0]
        T_Results['Volts'][byteloop*2 + 1]= roundNumber(self, hiFloat, base_precision_volt)

        ba = bytearray(struct.pack("d", DataI[byteloop]))   
        data = [ba[3], ba[2], ba[1], ba[0]]
        b = ''.join(chr(i) for i in data)
        lowFloat = struct.unpack('>f', b)[0]
        T_Results['Current'][byteloop*2]= roundNumber(self, lowFloat, base_precision_curr)
        data = [ba[7], ba[6], ba[5], ba[4]]
        b = ''.join(chr(i) for i in data)
        hiFloat = struct.unpack('>f', b)[0]
        T_Results['Current'][byteloop*2 + 1]= roundNumber(self, hiFloat, base_precision_curr)
        
    #Increment time_marker.
    for byteloop in range (0, samples):
        T_Results['Time'][byteloop] = roundNumber(self, time_marker, base_precision_time)
        time_marker = time_marker + sample_period
        
    T_Results['sample_count'] = samples 
    T_Results['sample_period'] = sample_period 
    T_Results['load_channel'] = load_num 
    T_Results['load_model'] = str_module_id  
    T_Results['load_frame'] = load_frame_id
    T_Results['dut_port'] = test_port  
#        T_Results['no_load_volt_post'] = self.roundNumber(after_volt, base_precision_volt)
    T_Results['no_load_volt_post'] = 0.5
       
    return T_Results, sample_period

def singlePortRamp(self, load_frame_id, test_port, load_num, current, duration, trigger_time):
    
    samples = self.max_samples
    cc_rise_slew_rate = None
    cc_fall_slew_rate = None

    os_time_on = None
    os_time_off = None
    load_on_seconds = None
    str_module_id = None 

    base_precision_volt = 1e-5 # precision for raw data
    base_precision_curr = 1e-5 # precision for raw data
    base_precision_time = 1e-6

    #T_Results is returned and becomes an instance of All_Results back in timerThread
    T_Results = newType_DataTable(self)
    #Add test code key to T_Results and set value
    T_Results['test_code'] = 1310
    
    current_range_margin_factor = 0.9
    
    #Add load_set_cc key and set value
    T_Results['load_set_cc'] = current 
	
    #Sample period = duration / max samples
    sample_period = float(duration) / samples # [2us~40ms, resolution 2us]
    sample_period = roundNumber(self, sample_period, self.sample_period_resolution) 
    self.sample_Period = sample_period
	
    if sample_period > self.max_sample_period:
        print("Sample period too long, setting to max: " + str(self.max_sample_period) + " seconds.") 
        sample_period = self.max_sample_period 
	
    #Recalculate duration because it may not match, rounding errors etc
    duration = samples * sample_period # calculate actual duration 
    trigger_sample = int(trigger_time / sample_period + 1) # trigger_sample range 1~4096 (or samples?)
	
    if trigger_sample > samples: 
        trigger_sample = samples 
	
    DataV = {}
    DataI = {} 

					
    # Common setup for all ports 					
    chromaLib.load_SelectSingleChannel(self.load, load_num) # select one load module  
    chromaLib.load_EnableSingleChannel(self.load, "ON") # [default ON] ??? Enable load module 
    chromaLib.load_TurnOnOffLoad(self.load, "OFF") # turn off load 
    chromaLib.load_TurnOnOffLoadShort(self.load, "OFF") #turn off short-circuit state  
    chromaLib.load_AllRunOnOff(self.load, "OFF") # [default ON] Turn off all load run (linked on/off) 
    chromaLib.load_AutoOnOff(self.load, "OFF") # [default OFF]  
    chromaLib.load_ConfigWindow(self.load, 0.10) # [default 0.02s] -- 2015-08-14 changed, only affects panel display 
    chromaLib.load_ConfigParallel(self.load, "NONE") # [default NONE] 
    chromaLib.load_ConfigSync(self.load, "NONE") # [default NONE] [Why set this?]  
    chromaLib.load_VoltLatchOnOff(self.load, "OFF") # [default OFF]
    chromaLib.load_VoltOff(self.load, 0.0) # [default 0.00] 
    chromaLib.load_VoltOn(self.load, 0.0) # [default 0.00] 
    chromaLib.load_VoltSign(self.load, "PLUS") # [default PLUS] 
	
    #This direct call to the load object should be in chromLib. Not sure why it's here?				
    str_module_id = self.load.query("CHAN:ID?")
    str_module_id = str_module_id.replace (',','_')
    str_module_id = str_module_id.replace ('\n','')
    str_module_id = str_module_id.replace ('\r','')
    
    # select left/right display for dual channel modules (NOTE: If both L&R used, then the last one set up is shown.) 
    if str_module_id.find("L_") != -1: 
        self.load.write("SHOW:DISP L") 
    elif str_module_id.find("R_") != -1:
        self.load.write("SHOW:DISP R") 
        
				
    LoadSpec = getLoadSpecs(self, str_module_id)
    max_ccm_amps        = LoadSpec['ccm_amps_max'] 
    max_ccl_amps        = LoadSpec['ccl_amps_max'] 
    
    if current > (max_ccm_amps * current_range_margin_factor):
        str_current_range = "OCPH"
    elif current > (max_ccl_amps * current_range_margin_factor):
        str_current_range = "OCPM"
    else:
        str_current_range = "OCPL"
                
    #Check Port IV range was here
    #Check Slew Rate was here
            
    print("Sampling rate = " + str(1 / sample_period) + " Hz for " + str(duration) + " seconds.") 

    chromaLib.load_SelectSingleChannel(self.load, load_num)
    chromaLib.load_CurrentMode(self.load, str_current_range)
    chromaLib.load_VoltRange(self.load, "LOW")

    print("Sampling rate = " + str(1 / sample_period) + " Hz for " + str(duration) + " seconds.") 
    print("Constant-current load = " + str(current) + " Amps, triggered at " + str((trigger_sample-1) * sample_period) + " seconds.")
     
    #Set Slew rates to max for now, not used in config spreadsheet right now
    chromaLib.load_RiseTime(self.load, "MAX") 
    chromaLib.load_FallTime(self.load, "MAX") 
    
    T_Results['top_axis'] = "current"
    #This stuff needs to be a function in chromaLib
    self.load.write("ADVANCE:OCP:ISTART 0")
    self.load.write("ADVANCE:OCP:IEND " + str(current))
    self.load.write("ADVANCE:OCP:STEP 1000")
    self.load.write("ADVANCE:OCP:DWELL 10ms")
    self.load.write("ADVANCE:OCP:TRIGGER:VOLTAGE 0.5")
				
    # Set up Digitizing mode 
    chromaLib.load_DigSampleTime(self.load, sample_period)  # [2us~40ms, resolution 2us] set sampling time  
    chromaLib.load_DigSamplePoints(self.load, samples) # [1~4096, default 4096] set sampling points 
    chromaLib.load_DigTriggerPoint(self.load, trigger_sample) # [1~4096, default 2000] set trigger point
    chromaLib.load_DigTriggerSource(self.load, "LOADON") # [default LOADON] set trigger source
    # Acquire V/I waveforms at load turn-on 
    self.load.write("DIG:INIT") # initialize digitizing 
 		
    while True:
        returnString = self.load.query("DIG:TRIG?") 
        if returnString.find('WAIT_TRIG') != -1:
            break
            
    os_time_on = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("Test started time: " + os_time_on) 

				
    chromaLib.load_TurnOnOffLoad(self.load, "ON")  
    
    while True:
        returnString = self.load.query("DIG:TRIG?") 
        if returnString.find("IDLE") != -1:
            break

    chromaLib.load_TurnOnOffLoad(self.load, "OFF")  
    
    os_time_off = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("Test finished time: " + os_time_off)
		
    # Read out waveforms  
    print("Reading Out Data.") 
				
    self.load.timeout = 30000
    while True:
        returnValue = self.load.query("DIG:WAV:CAP?")
        sleep(3)
        print returnValue
        if returnValue.find("OK") != -1:
            break
     
    DataV = self.load.query_binary_values("DIG:WAV:DATA? V", datatype='d', is_big_endian=False) # Largest buffer currently supported is 262144.
    DataI = self.load.query_binary_values("DIG:WAV:DATA? I", datatype='d', is_big_endian=False) # Largest buffer currently supported is 262144.

    T_Results['date_time'] = os_time_on
    T_Results['load_on_seconds'] = load_on_seconds
    T_Results['Volts'] = {}
    T_Results['Current'] = {}
    
    #Start time axis negative. At trigger point time will be zero
    time_marker = -sample_period * (trigger_sample - 1)
    
    for byteloop in range (0, samples/2):
        ba = bytearray(struct.pack("d", DataV[byteloop]))   
        data = [ba[3], ba[2], ba[1], ba[0]]
        b = ''.join(chr(i) for i in data)
        lowFloat = struct.unpack('>f', b)[0]
        T_Results['Volts'][byteloop*2]= roundNumber(self, lowFloat, base_precision_volt)
        data = [ba[7], ba[6], ba[5], ba[4]]
        b = ''.join(chr(i) for i in data)
        hiFloat = struct.unpack('>f', b)[0]
        T_Results['Volts'][byteloop*2 + 1]= roundNumber(self, hiFloat, base_precision_volt)

        ba = bytearray(struct.pack("d", DataI[byteloop]))   
        data = [ba[3], ba[2], ba[1], ba[0]]
        b = ''.join(chr(i) for i in data)
        lowFloat = struct.unpack('>f', b)[0]
        T_Results['Current'][byteloop*2]= roundNumber(self, lowFloat, base_precision_curr)
        data = [ba[7], ba[6], ba[5], ba[4]]
        b = ''.join(chr(i) for i in data)
        hiFloat = struct.unpack('>f', b)[0]
        T_Results['Current'][byteloop*2 + 1]= roundNumber(self, hiFloat, base_precision_curr)
        
    #Increment time_marker.
    for byteloop in range (0, samples):
        T_Results['Time'][byteloop] = roundNumber(self, time_marker, base_precision_time)
        time_marker = time_marker + sample_period
        
    T_Results['sample_count'] = samples 
    T_Results['sample_period'] = sample_period 
    T_Results['load_channel'] = load_num 
    T_Results['load_model'] = str_module_id  
    T_Results['load_frame'] = load_frame_id
    T_Results['dut_port'] = test_port  
#        T_Results['no_load_volt_post'] = self.roundNumber(after_volt, base_precision_volt)
    T_Results['no_load_volt_post'] = 0.5
       
    return T_Results, sample_period



def roundNumber(self, num1, factor1): 
    #num1 = number to round, factor1 = decimal precision (e.g. 0.01),  offset1 optional, default 0.5 (normal rounding) 
    result = None 
    offset1 = 0.5 
	
    if abs(num1) == num1:
        result = int(num1 / factor1 + offset1) * factor1 
    else: 
        result = int(num1 / factor1 - offset1) * factor1 

    return result  

def newType_DataTable(self):
    Type_DataTable = {} 
    Type_DataTable['Header'] = {} 
    Type_DataTable['Time'] = {}
    Type_DataTable['Volt'] = {}
    Type_DataTable['Curr'] = {} 
    return Type_DataTable 


def getLoadSpecs(self, str_load_module_id): 

    print "Module ID " + str_load_module_id
    ThisLoad = {}
    if str_load_module_id.find("63610-80-20") != -1: 
        ThisLoad['cch_amps_max'] = 20 
        ThisLoad['ccm_amps_max'] = 2 
        ThisLoad['ccl_amps_max'] = 0.2 
        ThisLoad['cch_slew_max'] = 2 # 2A/us 
        ThisLoad['cch_slew_min'] = 0.004 # 4A/ms 
        ThisLoad['ccm_slew_max'] = 0.2 # 0.2A/us
        ThisLoad['ccm_slew_min'] = 0.0004 # 0.4A/ms
        ThisLoad['ccl_slew_max'] = 0.02 # 0.02A/us
        ThisLoad['ccl_slew_min'] =	0.00004 # 0.04A/ms 
    elif str_load_module_id.find("63630-80-60") != -1:
        ThisLoad['cch_amps_max'] = 60 
        ThisLoad['ccm_amps_max'] = 6 
        ThisLoad['ccl_amps_max'] = 0.6 
        ThisLoad['cch_slew_max'] = 6 # 6A/us
        ThisLoad['cch_slew_min'] = 0.012 # 12A/ms
        ThisLoad['ccm_slew_max'] = 0.6 # 0.6A/us
        ThisLoad['ccm_slew_min'] = 0.0012 # 1.2A/ms
        ThisLoad['ccl_slew_max'] = 0.06 # 0.06A/us
        ThisLoad['ccl_slew_min'] =	0.00012 # 0.12A/ms 
    elif str_load_module_id.find("63640-80-80") != -1: 
        ThisLoad['cch_amps_max'] = 80 
        ThisLoad['ccm_amps_max'] = 8
        ThisLoad['ccl_amps_max'] = 0.8 
        ThisLoad['cch_slew_max'] = 8 # 8A/us
        ThisLoad['cch_slew_min'] = 0.016 # 16A/ms
        ThisLoad['ccm_slew_max'] = 0.8 # 0.8A/us
        ThisLoad['ccm_slew_min'] = 0.0016 # 1.6A/ms
        ThisLoad['ccl_slew_max'] = 0.08 # 0.08A/us
        ThisLoad['ccl_slew_min'] =	0.00016 # 0.16A/ms 
        
    return ThisLoad



