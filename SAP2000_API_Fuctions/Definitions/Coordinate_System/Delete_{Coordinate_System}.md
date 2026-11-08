# Delete

## Syntax

SapObject.SapModel.CoordSys.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing coordinate system.

## Remarks

The function deletes the specified coordinate system. Attempting to delete the Global coordinate system will fail and return an error.

The function returns zero if the coordinate system is successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteCoordSystem()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'define new coordinate system

      ret = SapModel.CoordSys.SetCoordSys("CSys1", 1000, 1000, 0, 0, 0, 0)



   'delete coordinate system

      ret = SapModel.CoordSys.Delete("CSys1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetCoordSys](SetCoordSys.md)
