# SetStageDefinitions (Note:  Newer Function Available)

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinear.SetStageDefinitions

## VB6 Procedure

Function SetStageDefinitions(ByVal Name As String, ByVal NumberStages As Long, ByRef Duration() As Long, ByRef Comment() As String) As Long

## Parameters

Name

The name of an existing static nonlinear staged analysis case.

NumberStages

The number of stages defined for the specified analysis case.

Duration

This is an array that includes the duration in days for each stage.

Comment

This is an array that includes a comment for each stage. The comment may be a blank string.

## Remarks

This function initializes the stage definition data for the specified load case. All previous stage definition data for the case is cleared when this function is called.

The function returns zero if the data is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseStaticNonlinearStagedStageDefinitions()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyDuration() As Long

      Dim MyComment() As String



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add static nonlinear staged load case

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetCase("LCASE1")



   'initialize stage definitions

      ReDim MyDuration(1)

      ReDim MyComment(1)

      MyDuration(0) = 0

      MyComment(0) = "Build structure"

      MyDuration(1) = 60

      MyComment(1) = "Wait"

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetStageDefinitions("LCASE1", 2, MyDuration, MyComment)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

This function is obsolete and has been superseded by [SetStageDefinitions\_1](SetStageDefinitions_1.md) as of version 12.00. This function is maintained for backwards compatibility.

## See Also

[GetStageDefinitions](GetStageDefinitions.md)
