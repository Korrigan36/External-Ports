import time

def load_SelectSingleChannel(loadObject, channel):
    tempString = "CHAN:LOAD " + str(channel) 
    loadObject.write(tempString)

def load_EnableSingleChannel(loadObject, text):
    tempString = "CHAN:ACT " + text 
    loadObject.write(tempString)

def load_TurnOnOffLoad(loadObject, text):
    tempString = "LOAD:STAT " + text 
    loadObject.write(tempString)

def load_TurnOnOffLoadShort(loadObject, text):
    tempString = "LOAD:SHOR " + text 
    loadObject.write(tempString)

def load_AllRunOnOff(loadObject, text):
    tempString = "CONF:ALLR " + text 
    loadObject.write(tempString)

def load_AutoOnOff(loadObject, text):
    tempString = "CONF:AUTO:ON " + text 
    loadObject.write(tempString)

def load_ConfigWindow(loadObject, wtime):
    tempString = "CONF:WIND " + str(wtime) 
    loadObject.write(tempString)

def load_ConfigParallel(loadObject, text):
    tempString = "CONF:PARA:MODE " + text 
    loadObject.write(tempString)

def load_ConfigSync(loadObject, text):
    tempString = "CONF:SYNC:MODE " + text 
    loadObject.write(tempString)

def load_VoltLatchOnOff(loadObject, text):
    tempString = "CONF:VOLT:LATC " + text 
    loadObject.write(tempString)

def load_VoltOff(loadObject, volts):
    tempString = "CONF:VOLT:OFF " + str(volts) 
    loadObject.write(tempString)

def load_VoltOn(loadObject, volts):
    tempString = "CONF:VOLT:ON " + str(volts) 
    loadObject.write(tempString)

def load_VoltSign(loadObject, text):
    tempString = "CONF:VOLT:SIGN " + text 
    loadObject.write(tempString)

def load_CurrentMode(loadObject, text):
    tempString = "MODE " + text 
    loadObject.write(tempString)

def load_VoltRange(loadObject, text):
    tempString = "CONF:VOLT:RANG " + text 
    loadObject.write(tempString)

def load_RiseTime(loadObject, text):
    tempString = "CURR:STAT:RISE " + text 
    loadObject.write(tempString)

def load_FallTime(loadObject, text):
    tempString = "CURR:STAT:FALL " + text 
    loadObject.write(tempString)

def load_DynamicRiseTime(loadObject, text):
    tempString = "CURR:DYN:RISE " + text 
    loadObject.write(tempString)

def load_DynamicFallTime(loadObject, text):
    tempString = "CURR:DYN:FALL " + text 
    loadObject.write(tempString)

def load_DynamicLoad1(loadObject, text):
    tempString = "CURR:DYN:L1 " + text 
    loadObject.write(tempString)

def load_DynamicLoad2(loadObject, text):
    tempString = "CURR:DYN:L2 " + text 
    loadObject.write(tempString)

def load_DynamicTime1(loadObject, text):
    tempString = "CURR:DYN:T1 " + text 
    loadObject.write(tempString)

def load_DynamicTime2(loadObject, text):
    tempString = "CURR:DYN:T2 " + text 
    loadObject.write(tempString)

def load_DynamicPulseCount(loadObject, text):
    tempString = "CURR:DYN:REP " + text 
    loadObject.write(tempString)

def load_SetCurrent(loadObject, current):
    tempString = "CURR:STAT:L1 " + str(current) 
    loadObject.write(tempString)

def load_DigSampleTime(loadObject, number):
    tempString = "DIG:SAMP:TIME " + str(number) 
    loadObject.write(tempString)

def load_DigSamplePoints(loadObject, number):
    tempString = "DIG:SAMP:POIN " + str(number) 
    loadObject.write(tempString)

def load_DigTriggerPoint(loadObject, number):
    tempString = "DIG:TRIG:POIN " + str(number) 
    loadObject.write(tempString)

def load_DigTriggerSource(loadObject, text):
    tempString = "DIG:TRIG:SOUR " + text 
    loadObject.write(tempString)

def load_SetActiveOff(loadObject, text):
    tempString = "CHAN:LOAD " + text 
    loadObject.write(tempString)
    tempString = "CHAN:ACT OFF"
    loadObject.write(tempString)
    
#def load_TurnOffParallelMode(loadObject):
#    tempString = "CONF:PARA:INIT OFF"
#    loadObject.write(tempString)
#
#


