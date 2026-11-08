# GetCaseStatus

## Syntax

SapObject.SapModel.Analyze.GetCaseStatus

## VB6 Procedure

Function GetCaseStatus(ByRef NumberItems As Long, ByRef CaseName() As String, ByRef Status() As Long) As Long

## Parameters

NumberItems

The number of load cases for which the status is reported.

CaseName

This is an array that includes the name of each analysis case for which the status is reported.

Status

This is an array of that includes 1, 2, 3 or 4, indicating the load case status.

1 = Not run

2 = Could not start

3 = Not finished

4 = Finished

## Remarks

This function retrieves the status for all load cases.

The function returns zero if the status is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStatusExample()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim CaseName() As String

      Dim Status() As Long



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



   'run model (this will create the analysis model)

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'get load case status

      ret = SapModel.Analyze.GetCaseStatus(NumberItems, CaseName, Status)



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
