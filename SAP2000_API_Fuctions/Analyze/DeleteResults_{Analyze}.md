# DeleteResults

## Syntax

SapObject.SapModel.Analyze.DeleteResults

## VB6 Procedure

Function DeleteResults(ByVal Name As String, Optional ByVal All As Boolean = False) As Long

## Parameters

Name

The name of an existing load case that is to have its results deleted.

This item is ignored when the All item is True.

All

If this item is True, the results are deleted for all load cases, and the Name item is ignored.

## Remarks

This function deletes results for load cases.

The function returns zero if the results are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteLoadCaseResults()

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



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'delete load case results

      ret = SapModel.Analyze.DeleteResults("MODAL")



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
