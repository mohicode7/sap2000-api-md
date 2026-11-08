# GetPoints

## Syntax

SapObject.SapModel.FrameObj.GetPoints

## VB6 Procedure

Function GetPoints(ByVal Name As String, ByRef Point1 As String, ByRef Point2 As String) As Long

## Parameters

Name

The name of a defined frame object.

Point1

The name of the point object at the I-End of the specified frame object.

Point2

The name of the point object at the J-End of the specified frame object.

## Remarks

This function retrieves the names of the point objects at each end of a specified frame object.

The function returns zero if the point names are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameObjPoints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Point1 As String

      Dim Point2 As String



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



   'get names of points

      ret = SapModel.FrameObj.GetPoints("3", Point1, Point2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
