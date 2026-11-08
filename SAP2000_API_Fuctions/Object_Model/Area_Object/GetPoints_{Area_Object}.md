# GetPoints

## Syntax

SapObject.SapModel.AreaObj.GetPoints

## VB6 Procedure

Function GetPoints(ByVal Name As String, ByRef NumberPoints As Long, ByRef Point() As String) As Long

## Parameters

Name

The name of a defined area object.

NumberPoints

The number of point objects that define the area object.

Point

This is an array containing the names of the point objects that define the area object. The point names are in order around the area object.

## Remarks

This function retrieves the names of the point objects that define an area object.

The function returns zero if the point object names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjPoints()

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



   'get names of points

      ret = SapModel.AreaObj.GetPoints("1", NumberPoints, Point)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
