# SetStageDefinitions\_2

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.SetStageDefinitions\_2

## VB6 Procedure

Function SetStageDefinitions\_2(ByVal Name As String,
ByVal NumberStages As Long, ByRef Duration() As Double, ByRef Output()
As Boolean, ByRef OutputName() As String, ByRef Comment() As String) As
Long

## Parameters

Name

The name of an existing static nonlinear staged load
case.

NumberStages

The number of stages defined for the specified load
case.

Duration

This is an array that includes the duration in days
for each stage.

Output

This is an array that includes True or False, indicating
if analysis output is to be saved for each stage.

OutputName

This is an array that includes a user-specified output
name for each stage.

Comment

This is an array that includes a comment for each stage.
The comment may be a blank string.

## Remarks

This function initializes the stage definition data
for the specified load case. All previous stage definition data for the
case is cleared when this function is called.

The function returns zero if the data is successfully
initialized; otherwise, it returns a nonzero value.

## VBA Example



```
Sub SetCaseStaticNonlinearStagedStageDefinitions_2()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyDuration() As Double

      Dim MyOutput() As Boolean

      Dim MyOutputName() As String

      Dim MyComment() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'add static nonlinear staged load case

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetCase("ACASE1")



   'initialize stage definitions

      ReDim MyDuration(1)

      ReDim MyOutput(1)

      ReDim MyOutputName(1)

      ReDim MyComment(1)

      MyDuration(0) = 0

      MyOutput(0) = False

      MyComment(0) = "Build structure"

      MyDuration(1) = 60

      MyOutput(1) = True

      MyOutputName(1) = "HBC2"

      MyComment(1) = "Wait"

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetStageDefinitions_2("ACASE1", 
 2, MyDuration, MyOutput, MyOutputName, MyComment)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

This function supersedes [GetStageDefinitions\_1](../../../Obsolete_Functions/GetStageDefinitions_1.md).

## See Also

[GetStageDefinitions\_2](GetStageDefinitions_2.md)
