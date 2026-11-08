# GetPresentCoordSystem

## Syntax

SapObject.SapModel.GetPresentCoordSystem

## VB6 Procedure

Function GetPresentCoordSystem() As String

## Parameters

None

## Remarks

This function returns the name of the present coordinate system.

## VBA Example



```
Sub GetPresentCSys()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as long

      Dim PresentCSys as String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(ConcentricBraced, 3, 124, 3, 200)



   'define new coordinate system

      ret = SapModel.CoordSys.SetCoordSys("CSys1", 1000, 1000, 0, 0, 0, 0)



   'set present coordinate system

      ret = SapModel.SetPresentCoordSystem("CSys1")



   'get present coordinate system

      PresentCSys = SapModel.GetPresentCoordSystem



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetPresentCoordSystem](SetPresentCoordSystem.md)
