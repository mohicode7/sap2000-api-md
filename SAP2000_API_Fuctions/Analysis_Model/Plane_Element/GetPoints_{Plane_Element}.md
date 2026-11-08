# GetPoints

## Syntax

SapObject.SapModel.PlaneElm.GetPoints

## VB6 Procedure

Function GetPoints(ByVal Name As String, ByRef NumberPoints As Long, ByRef Point() As String) As Long

## Parameters

Name

The name of an plane element.

NumberPoints

The number of point elements that define the plane element.

Point

This is an array containing the names of the point elements that define the plane element. The point names are in order around the plane element.

## Remarks

This function retrieves the names of the point elements that define an plane element.

The function returns zero if the point element names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPlaneElmPoints()

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



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 3-001-incomp.sdb")



   'assign auto mesh options

      ret = SapModel.AreaObj.SetAutoMesh("ALL", 1, 3, 3, , , , , , , , , , , , , , , Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get names of points

      ret = SapModel.PlaneElm.GetPoints("3-2", NumberPoints, Point)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
