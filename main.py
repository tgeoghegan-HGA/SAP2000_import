import os
import sys
import comtypes.client
from HGAtools import *

printhello()
#set the following flag to True to attach to an existing instance of the program
#otherwise a new instance of the program will be started
AttachToInstance = True

#set the following flag to True to manually specify the path to SAP2000.exe
#this allows for a connection to a version of SAP2000 other than the latest installation
#otherwise the latest installed version of SAP2000 will be launched
SpecifyPath = False


#create API helper object
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject") 


if AttachToInstance:
    #attach to a running instance of SAP2000
    try:
        #get the active SapObject
            mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject") 


    except (OSError, comtypes.COMError):

        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)


SapModel = mySapObject.SapModel
'''
GetTableForDisplayArray
GetTableForDisplayCSVFile
GetTableForDisplayCSVString
GetTableForDisplayXMLString
GetTableForEditingArray
GetTableForEditingCSVFile
GetTableForEditingCSVString
GetTableOutputOptionsForDisplay
'''

Get_table_E('Connectivity - Frame', SapModel)
