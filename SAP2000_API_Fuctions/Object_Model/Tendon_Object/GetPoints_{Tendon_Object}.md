# GetPoints

## Syntax

SapObject.SapModel.TendonObj.GetPoints

## VB6 Procedure

Function GetPoints(ByVal Name As String, ByRef Point1 As String, ByRef Point2 As String) As Long

## Parameters

Name

The name of a defined tendon object.

Point1

The name of the point object at the I-End of the specified tendon object.

Point2

The name of the point object at the J-End of the specified tendon object.

## Remarks

This function retrieves the names of the point objects at each end of a specified tendon object.

The function returns zero if the point names are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonObjPoints()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon object by coordinates

      ret = SapModel.TendonObj.AddByCoord(-288, 0, 288, 288, 0, 288, Name)



   'get names of points

      ret = SapModel.TendonObj.GetPoints(Name, Point1, Point2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
