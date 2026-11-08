# SetRunCaseFlag

## Syntax

SapObject.SapModel.Analyze.SetRunCaseFlag

## VB6 Procedure

Function SetRunCaseFlag(ByVal Name As String, ByVal Run As Boolean, Optional ByVal All As Boolean = False) As Long

## Parameters

Name

The name of an existing load case that is to have its run flag set.

This item is ignored when the All item is True.

Run

If this item is True, the specified load case is to be run.

All

If this item is True, the run flag is set as specified by the Run item for all load cases, and the Name item is ignored.

## Remarks

This function sets the run flag for load cases.

The function returns zero if the flag is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRunCaseFlagExample()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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

[GetRunCaseFlag](GetRunCaseFlag.md)
