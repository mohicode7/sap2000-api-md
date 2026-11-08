# GetRunCaseFlag

## Syntax

SapObject.SapModel.Analyze.GetRunCaseFlag

## VB6 Procedure

Function GetRunCaseFlag(ByRef NumberItems As Long, ByRef CaseName() As String, ByRef Run() As Boolean) As Long

## Parameters

NumberItems

The number of load cases for which the run flag is reported.

CaseName

This is an array that includes the name of each analysis case for which the run flag is reported.

Run

This is an array of boolean values indicating if the specified load case is to be run.

## Remarks

This function retrieves the run flags for all analysis cases.

The function returns zero if the flags are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRunCaseFlagExample()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim CaseName() As String

      Dim Run() As Boolean



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set load case run flag

      ret = SapModel.Analyze.SetRunCaseFlag("MODAL", False)



   'get load case run flags

      ret = SapModel.Analyze.GetRunCaseFlag(NumberItems, CaseName, Run)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SetRunCaseFlag](SetRunCaseFlag.md)
