# Example  (IronPython)

## Remarks

This example was created using IronPython 3.4 . It is based on the SAP2000 verification problem Example 1-001.

This example creates the example verification problem from scratch, runs the analysis, extracts results, and compares the results with hand calculated values.

## Example

1. Either download and install the IronPython interpreter from https:\\ironpython.net or, on a command prompt, navigate to your API script folder and install IronPython as a .NET local tool by issuing the following commands:

dotnet new tool-manifest

dotnet tool install --local IronPython.Console --version 3.4.1

2. Create a Python .py file using your preferred IDE, copy and paste in the Python code example below, and save it as IronPythonApplication1.py in your API script folder.
3. When using the IronPython interpreter, open a command prompt, navigate to your API script folder and issue the following command to run the API script

When using the IronPython interpreter, open a command prompt, navigate to your API script folder and issue the following command to run the API script:



```
ipy IronPythonApplication1.py

```



When using the .NET local tool, on the same command prompt, issue the following command to run the API script:



```
dotnet ipy IronPythonApplication1.py

```



Please pay attention to the comments in this code block; they contain important information about running the script. This example includes an option for executing on a remote computer, however that functionality has been disabled in SAP2000 v26.0.0. It may be reinstated in a future release.



```
import os
import sys
import clr


from enum import Enum


clr.AddReference("System")


import System

#set the following path to the installed SAP2000 program directory
clr.AddReferenceToFileAndPath("C:\Program Files\Computers and Structures\SAP2000 26\SAP2000v1.dll")


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
ProgramPath = "C:\Program Files\Computers and Structures\SAP2000 26\SAP2000.exe"


#full path to the model
#set it to the desired path of your model
#if executing remotely, ensure that this folder already exists on the remote computer
#the below command will only create the folder locally
APIPath = "C:\CSiAPIExample"
if not os.path.exists(APIPath):
    try:
        os.makedirs(APIPath)
    except OSError:
        pass

ModelPath = APIPath + os.sep + "API_1-001.sdb"


#create API helper object
helper = Helper()


if AttachToInstance:
    #attach to a running instance of SAP2000
    try:
        #get the active SapObject      
        if Remote:
            mySapObject = helper.GetObjectHost(RemoteComputer, "CSI.SAP2000.API.SapObject")
        else:
            mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject")
    except:
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)
else:
    if SpecifyPath:
        try:
            #'create an instance of the SapObject from the specified path
            if Remote:
                mySapObject = helper.CreateObjectHost(RemoteComputer, ProgramPath)
            else:
                mySapObject = helper.CreateObject(ProgramPath)
        except:
            print("Cannot start a new instance of the program from " + ProgramPath)
            sys.exit(-1)
    else:
        try:
            #create an instance of the SapObject from the latest installed SAP2000
            if Remote:
                mySapObject = helper.CreateObjectProgIDHost(RemoteComputer, "CSI.SAP2000.API.SapObject")
            else:
                mySapObject = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
        except:
            print("Cannot start a new instance of the program.")
            sys.exit(-1)


    #start SAP2000 application
    mySapObject.ApplicationStart()


#create SapModel object
SapModel = mySapObject.SapModel


#initialize model
SapModel.InitializeNewModel()


#create new blank model
File = SapModel.File
ret = File.NewBlank()


#define material property
PropMaterial = SapModel.PropMaterial
ret = PropMaterial.SetMaterial('CONC', eMatType.Concrete)


#assign isotropic mechanical properties to material
ret = PropMaterial.SetMPIsotropic('CONC', 3600, 0.2, 0.0000055)


#define rectangular frame section property
PropFrame = SapModel.PropFrame
ret = PropFrame.SetRectangle('R1', 'CONC', 12, 12)


#define frame section property modifiers
ModValue = System.Array[float]([1000, 0, 0, 1, 1, 1, 1, 1])
ret = PropFrame.SetModifiers('R1', ModValue)


#switch to k-ft units
ret = SapModel.SetPresentUnits(eUnits.kip_ft_F)


#add frame object by coordinates
FrameObj = SapModel.FrameObj
FrameName1 = ' '
FrameName2 = ' '
FrameName3 = ' '

[ret, FrameName1] = FrameObj.AddByCoord(0, 0, 0, 0, 0, 10, FrameName1, 'R1', '1', 'Global')
[ret, FrameName2] = FrameObj.AddByCoord(0, 0, 10, 8, 0, 16, FrameName2, 'R1', '2', 'Global')
[ret, FrameName3] = FrameObj.AddByCoord(-4, 0, 10, 0, 0, 10, FrameName3, 'R1', '3', 'Global')


#assign point object restraint at base
PointObj = SapModel.PointObj
PointName1 = ' '
PointName2 = ' '
Restraint = clr.StrongBox[System.Array[bool]](System.Array[bool]([True, True, True, True, False, False]))
[ret, PointName1, PointName2] = FrameObj.GetPoints(FrameName1, PointName1, PointName2)
ret = PointObj.SetRestraint(PointName1, Restraint)


#assign point object restraint at top
Restraint = clr.StrongBox[System.Array[bool]](System.Array[bool]([True, True, False, False, False, False]))
[ret, PointName1, PointName2] = FrameObj.GetPoints(FrameName2, PointName1, PointName2)
ret = PointObj.SetRestraint(PointName2, Restraint)


#refresh view, update (initialize) zoom
View = SapModel.View
ret = View.RefreshView(0, False)


#add load patterns
LoadPatterns = SapModel.LoadPatterns
ret = LoadPatterns.Add('1', eLoadPatternType.Other, 1, True)
ret = LoadPatterns.Add('2', eLoadPatternType.Other, 0, True)
ret = LoadPatterns.Add('3', eLoadPatternType.Other, 0, True)
ret = LoadPatterns.Add('4', eLoadPatternType.Other, 0, True)
ret = LoadPatterns.Add('5', eLoadPatternType.Other, 0, True)
ret = LoadPatterns.Add('6', eLoadPatternType.Other, 0, True)
ret = LoadPatterns.Add('7', eLoadPatternType.Other, 0, True)


#assign loading for load pattern 2
[ret, PointName1, PointName2] = FrameObj.GetPoints(FrameName3, PointName1, PointName2)
PointLoadValue = clr.StrongBox[System.Array[float]](System.Array[float]([0,0,-10,0,0,0]))
ret = PointObj.SetLoadForce(PointName1, '2', PointLoadValue)
ret = FrameObj.SetLoadDistributed(FrameName3, '2', 1, 10, 0, 1, 1.8, 1.8)


#assign loading for load pattern 3
[ret, PointName1, PointName2] = FrameObj.GetPoints(FrameName3, PointName1, PointName2)
PointLoadValue = clr.StrongBox[System.Array[float]](System.Array[float]([0,0,-17.2,0,-54.4,0]))
ret = PointObj.SetLoadForce(PointName2, '3', PointLoadValue)


#assign loading for load pattern 4
ret = FrameObj.SetLoadDistributed(FrameName2, '4', 1, 11, 0, 1, 2, 2)


#assign loading for load pattern 5
ret = FrameObj.SetLoadDistributed(FrameName1, '5', 1, 2, 0, 1, 2, 2, 'Local')
ret = FrameObj.SetLoadDistributed(FrameName2, '5', 1, 2, 0, 1, -2, -2, 'Local')


#assign loading for load pattern 6
ret = FrameObj.SetLoadDistributed(FrameName1, '6', 1, 2, 0, 1, 0.9984, 0.3744, 'Local')
ret = FrameObj.SetLoadDistributed(FrameName2, '6', 1, 2, 0, 1, -0.3744, 0, 'Local')


#assign loading for load pattern 7
ret = FrameObj.SetLoadPoint(FrameName2, '7', 1, 2, 0.5, -15, 'Local')


#switch to k-in units
ret = SapModel.SetPresentUnits(eUnits.kip_in_F)


#save model
File = SapModel.File
ret = File.Save(ModelPath)


#run model (this will create the analysis model)
Analyze = SapModel.Analyze
ret = Analyze.RunAnalysis()


#initialize for results
ProgramResult = System.Array[float]([0,0,0,0,0,0,0])
[ret, PointName1, PointName2] = FrameObj.GetPoints(FrameName2, PointName1, PointName2)


#get results for load cases 1 through 7
Results = SapModel.Results
Setup = Results.Setup
for i in range(0,7):
      NumberResults = 0
      Obj = System.Array[str]([])
      Elm = System.Array[str]([])
      ACase = System.Array[str]([])
      StepType = System.Array[str]([])
      StepNum = System.Array[float]([])
      U1 = System.Array[float]([])
      U2 = System.Array[float]([])
      U3 = System.Array[float]([])
      R1 = System.Array[float]([])
      R2 = System.Array[float]([])
      R3 = System.Array[float]([])
      ObjectElm = 0
      ret = Setup.DeselectAllCasesAndCombosForOutput()
      ret = Setup.SetCaseSelectedForOutput(str(i + 1))

      if i <= 3:
          [ret, NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3] = Results.JointDispl(PointName2, eItemTypeElm(ObjectElm), NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
          ProgramResult[i] = U3[0]
      else:
          [ret, NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3] = Results.JointDispl(PointName1, eItemTypeElm(ObjectElm), NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
          ProgramResult[i] = U1[0]


#close the program
ret = mySapObject.ApplicationExit(False)
SapModel = None
mySapObject = None


#fill independent results
IndResult = [0,0,0,0,0,0,0]
IndResult[0] = -0.02639
IndResult[1] = 0.06296
IndResult[2] = 0.06296
IndResult[3] = -0.2963
IndResult[4] = 0.3125
IndResult[5] = 0.11556
IndResult[6] = 0.00651


#fill percent difference
PercentDiff = [0,0,0,0,0,0,0]
for i in range(0,7):
      PercentDiff[i] = (ProgramResult[i] / IndResult[i]) - 1


#display results
for i in range(0,7):
      print()
      print(ProgramResult[i])
      print(IndResult[i])
      print(PercentDiff[i])
```



## Release Notes

Initial Release for version 24.2.0

Updated for version 26.0.0.
