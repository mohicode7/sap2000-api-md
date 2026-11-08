# ChangeConnectivity

## Syntax

SapObject.SapModel.EditArea.ChangeConnectivity

## VB6 Procedure

Function ChangeConnectivity(ByVal Name As String, ByVal NumberPoints As Long, ByRef Point() As String) As Long

## Parameters

Name

The name of an existing area object.

NumberPoints

The number of points in the area abject.

Point

This is an array containing the names of the point objects that define the added area object. The point object names should be ordered to run clockwise or counter-clockwise around the area object.

## Remarks

This function modifies the connectivity of an area object.

The function returns zero if the area object connectivity is successfully modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub ModifyAreaObjConnectivity()

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

      ret = SapModel.AreaObj.GetPoints("2", NumberPoints, Point)



   'modify connectivity

      NumberPoints = NumberPoints - 1

      ReDim Preserve Point(NumberPoints )

      ret = SapModel.EditArea.ChangeConnectivity("2", NumberPoints, Point)

      ret = SapModel.View.RefreshWindow



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also
