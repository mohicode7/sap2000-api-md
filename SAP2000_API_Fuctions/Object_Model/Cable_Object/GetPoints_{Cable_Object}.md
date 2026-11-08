# GetPoints

## Syntax

SapObject.SapModel.CableObj.GetPoints

## VB6 Procedure

Function GetPoints(ByVal Name As String, ByRef Point1 As String, ByRef Point2 As String) As Long

## Parameters

Name

The name of a defined cable object.

Point1

The name of the point object at the I-End of the specified cable object.

Point2

The name of the point object at the J-End of the specified cable object.

## Remarks

This function retrieves the names of the point objects at each end of a specified cable object.

The function returns zero if the point names are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableObjPoints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Point1 As String

      Dim Point2 As String

      Dim Name As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add cable object by coordinates

      ret = SapModel.CableObj.AddByCoord(-300, 0, 0, -100, 0, 124, Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'get names of points

      ret = SapModel.CableObj.GetPoints(Name, Point1, Point2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
