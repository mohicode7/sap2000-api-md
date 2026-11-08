# GetSpecialPoint

## Syntax

SapObject.SapModel.PointObj.GetSpecialPoint

## VB6 Procedure

Function GetSpecialPoint(ByVal Name As String, ByVal SpecialPoint As Boolean) As Long

## Parameters

Name

The name of an existing point object.

SpecialPoint

This item is True if the point object is specified as a special point, otherwise it is False.

## Remarks

This function retrieves the special point status for a point object.

The function returns zero if the special point status is successfully retrieved, otherwise it returns a nonzero value.

Special points are allowed to exist in the model even if no objects (line, area, solid, link) are connected to them. Points that are not special are automatically deleted if no objects connect to them.

## VBA Example



```
Sub GetSpecialPointStatus()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SpecialPoint As Boolean



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



   'set as special point

      ret = SapModel.PointObj.SetSpecialPoint("3", True)



   'get special point status

      ret = SapModel.PointObj.GetSpecialPoint("3", SpecialPoint)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetSpecialPoint](SetSpecialPoint.md)

[DeleteSpecialPoint](DeleteSpecialPoint.md)
