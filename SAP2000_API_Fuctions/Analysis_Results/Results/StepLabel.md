# StepLabel

## Syntax

SapObject.SapModel.Results.StepLabel

## VB6 Procedure

Function StepLabel(ByVal LoadCase As String, ByVal StepNum As Double, ByRef Label As String) As Long

## Parameters

LoadCase

The name of an existing linear multi-step, nonlinear multi-step, or staged-construction load case.

StepNum

This is an overall step number from the specified load case. The range of values of StepNum for a given load case can be obtained from most analysis results calls, such as SapObject.SapModel.Results.JointDispl. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

Label

The is the step label, including the name or number of the stage, the step number within the stage, and the age of the structure for time-dependent load cases

## Remarks

This function generates the step label for analyzed linear multi-step, nonlinear multi-step, or staged-construction load cases. For other load case types, the label will be blank.

The function returns zero if the step label is successfully generated; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetStepLabel()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As Sap2000v16.cSapModel

          Dim ret As Long

          Dim MyDuration() As Long

          Dim MyOutput() As Boolean

          Dim MyOutputName() As String

          Dim MyComment() As String

          Dim MyOperation() As Long

          Dim MyObjectType() As String

          Dim MyObjectName() As String

          Dim MyAge() As Long

          Dim MyMyType() As String

          Dim MyMyName() As String

          Dim MySF() As Double

          Dim Label() as String

     
   'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'create model from template

          ret = SapModel.File.New2DFrame(Sap2000v16.PortalFrame, 2, 144, 2, 288)



       'add static nonlinear staged load case

          ret = SapModel.LoadCases.StaticNonlinearStaged.SetCase("ACASE1")



       'initialize stage definitions

          ReDim MyDuration(1)

          ReDim MyOutput(1)

          ReDim MyOutputName(1)

          ReDim MyComment(1)
  
      Redim Label(1)

          MyDuration(0) = 0

          MyOutput(0) = False

          MyComment(0) = "Build structure"

          MyDuration(1) = 60

          MyOutput(1) = True

          MyOutputName(1) = "HBC2"

          MyComment(1) = "Wait"

          ret = SapModel.LoadCases.StaticNonlinearStaged.SetStageDefinitions_1("ACASE1", 2, MyDuration, MyOutput, MyOutputName, MyComment)



       'set stage data

          ReDim MyOperation(1)

          ReDim MyObjectType(1)

          ReDim MyObjectName(1)

          ReDim MyAge(1)

          ReDim MyMyType(1)

          ReDim MyMyName(1)

          ReDim MySF(1)

          MyOperation(0) = 1

          MyObjectType(0) = "Group"

          MyObjectName(0) = "ALL"

          MyAge(0) = 3

          MyOperation(1) = 4

          MyObjectType(1) = "Frame"

          MyObjectName(1) = "8"

          MyMyType(1) = "Load"

          MyMyName(1) = "DEAD"

          MySF(1) = 0.85

          ret = SapModel.LoadCases.StaticNonlinearStaged.SetStageData_1("ACASE1", 1, 2, MyOperation, MyObjectType, MyObjectName, MyAge, MyMyType, MyMyName, MySF)



       'set results saved parameters

          ret = SapModel.LoadCases.StaticNonlinearStaged.SetResultsSaved("ACASE1", 3, 4, 10)



       'save model

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")



       'run model (this will create the analysis model)

          ret = SapModel.Analyze.RunAnalysis



       'Get step label

          ret = SapModel.Results.StepLabel("ACASE1", 3, Label(0))

          ret = SapModel.Results.StepLabel("ACASE1", 8, Label(1))



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 16.00.
