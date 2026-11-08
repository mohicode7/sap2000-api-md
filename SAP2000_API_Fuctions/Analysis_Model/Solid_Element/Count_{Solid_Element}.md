# Count

## Syntax

Sap2000.SolidElm.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns the total number of solid elements in the analysis model.

## VBA Example



```
Sub CountSolidElements()

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



   'create model from template

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'assign auto mesh options

      ret = SapModel.SolidObj.SetAutoMesh("ALL", 1, 2, 2, 2, , , , , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'return number of solid elements

      Count = SapModel.SolidElm.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
