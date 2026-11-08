# GetPoints

## Syntax

SapObject.SapModel.AreaElm.GetPoints

## VB6 Procedure

Function GetPoints(ByVal Name As String, ByRef NumberPoints As Long, ByRef Point() As String) As Long

## Parameters

Name

The name of an area element.

NumberPoints

The number of point elements that define the area element.

Point

This is an array containing the names of the point elements that define the area element. The point names are in order around the area element.

## Remarks

This function retrieves the names of the point elements that define an area element.

The function returns zero if the point element names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaElmPoints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberPoints As Long

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'assign auto mesh options

      ret = SapModel.AreaObj.SetAutoMesh("ALL", 1, 3, 3, , , , , , , , , , , , , , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get names of points

      ret = SapModel.AreaElm.GetPoints("3-2", NumberPoints, Point)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
