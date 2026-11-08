# Launching the Installed Version of SAP2000/CSiBridge Automatically

With the release of SAP2000 and CSiBridge v19.1.0, new functionality has been added to the cHelper interface to allow users to launch the application without supplying the path to the program executable file. The example code is in Visual Basic.

1. Previously, users would launch the program with the function cHelper.CreateObject , providing the full path to the SAP2000.exe as an argument, as below:

Dim mySapObject As SAP2000v19.cOAPI

Dim myHelper as SAP2000v19.cHelper = New SAP2000v19.Helper

mySapObject = myHelper.CreateObject("C:\Program Files (x86)\Computers and Structures\SAP2000 19\SAP2000.exe")

ret = mySapObject.ApplicationStart()

Dim mySapModel As SAP2000v19.cSapModel = mySapObject.SapModel

2. While the code above is still available, a simpler method has been added to cHelper. The CreateObjectProgID function takes the Program ID as an argument, and automatically launches the most recently installed version of SAP2000 or CSiBridge:

Dim mySapObject As SAP2000v19.cOAPI

Dim myHelper as SAP2000v19.cHelper = New SAP2000v19.Helper

mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject")

ret = mySapObject.ApplicationStart()

Dim mySapModel As SAP2000v19.cSapModel = mySapObject.SapModel

For reference, the Program ID for CSiBridge users is "CSI.CSiBridge.API.SapObject"
