# ChangeName

## Syntax

SapObject.SapModel.CoordSys.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined coordinate system.

NewName

The new name for the coordinate system.

## Remarks

The function returns zero if the new name is successfully applied, otherwise it returns a nonzero value.

Changing the name of the Global coordinate system will fail and return an error.

## VBA Example



```
Sub ChangeCoordSystemName()

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



   'change name of new coordinate system

      ret = SapModel.CoordSys.ChangeName("CSys1", "MyCSys")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
