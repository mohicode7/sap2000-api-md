# GetPoints

## Syntax

SapObject.SapModel.SolidObj.GetPoints

## VB6 Procedure

Function GetPoints(ByVal Name As String, ByRef Point() As String) As Long

## Parameters

Name

The name of a defined solid object.

Point

This is an array containing the names of the corner point objects of the solid object.

## Remarks

This function retrieves the names of the corner point objects of a solid object.

The function returns zero if the point object names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidObjPoints()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'get names of points

      ReDim Point(7)

      ret = SapModel.SolidObj.GetPoints("1", Point)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
