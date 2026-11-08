# GetPoints

## Syntax

SapObject.SapModel.SolidElm.GetPoints

## VB6 Procedure

Function GetPoints(ByVal Name As String, ByRef Point() As String) As Long

## Parameters

Name

The name of an solid element.

Point

This is an array containing the names of the eight point elements that define the solid element. The point names are in order around the solid element.

## Remarks

This function retrieves the names of the eight point elements that define a solid element.

The function returns zero if the point element names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidElmPoints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Point() As String



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



   'get names of eight corner points

      ret = SapModel.SolidElm.GetPoints("3-2", Point)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
