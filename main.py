import os
import sys

import clr
 
from enum import Enum

clr.AddReference("System.Runtime.InteropServices")
from System.Runtime.InteropServices import Marshal

#set the following path to the installed SAP2000 program directory
clr.AddReference(R'C:\Program Files\Computers and Structures\SAP2000 24\SAP2000v1.dll')

from SAP2000v1 import *

 

#set the following flag to True to execute on a remote computer

Remote = False

 

#if the above flag is True, set the following variable to the hostname of the remote computer

#remember that the remote computer must have SAP2000 installed and be running the CSiAPIService.exe

RemoteComputer = "SpareComputer-DT"

 

#set the following flag to True to attach to an existing instance of the program

#otherwise a new instance of the program will be started

AttachToInstance = False

 

#set the following flag to True to manually specify the path to SAP2000.exe

#this allows for a connection to a version of SAP2000 other than the latest installation

#otherwise the latest installed version of SAP2000 will be launched

SpecifyPath = False

 

#if the above flag is set to True, specify the path to SAP2000 below

ProgramPath = R"C:\Program Files\Computers and Structures\SAP2000 24\SAP2000.exe"

 

#full path to the model

#set it to the desired path of your model

#if executing remotely, ensure that this folder already exists on the remote computer

#the below command will only create the folder locally

APIPath = R'C:\CSi_SAP2000_API_Example'

if not os.path.exists(APIPath):

    try:

        os.makedirs(APIPath)

    except OSError:

        pass

ModelPath = APIPath + os.sep + 'API_1-001.sdb'

 

#create API helper object

helper = cHelper(Helper())

 

if AttachToInstance:

    #attach to a running instance of SAP2000

    try:

        #get the active SAP2000 object       

        if Remote:

            mySAPObject = cOAPI(helper.GetObjectHost(RemoteComputer, "CSI.SAP2000.API.SAPObject"))

        else:

            mySAPObject = cOAPI(helper.GetObject("CSI.SAP2000.API.SAPObject"))

    except:

        print("No running instance of the program found or failed to attach.")

        sys.exit(-1)
