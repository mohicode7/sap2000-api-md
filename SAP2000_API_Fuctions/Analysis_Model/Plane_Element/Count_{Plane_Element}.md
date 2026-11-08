# Count

## Syntax

Sap2000.PlaneElm.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns the total number of plane elements in the analysis model.

## VBA Example



```
Sub CountPlaneElements()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 3-001-incomp.sdb")



   'assign auto mesh options

      ret = SapModel.AreaObj.SetAutoMesh("ALL", 1, 2, 2, , , , , , , , , , , , , , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'return number of plane elements

      Count = SapModel.PlaneElm.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
