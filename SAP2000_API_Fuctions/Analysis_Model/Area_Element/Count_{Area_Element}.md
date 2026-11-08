# Count

## Syntax

Sap2000.AreaElm.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns the total number of area elements in the analysis model.

## VBA Example



```
Sub CountAreaElements()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'assign auto mesh options

      ret = SapModel.AreaObj.SetAutoMesh("ALL", 1, 3, 3, , , , , , , , , , , , , , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'return number of area elements

      Count = SapModel.AreaElm.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
