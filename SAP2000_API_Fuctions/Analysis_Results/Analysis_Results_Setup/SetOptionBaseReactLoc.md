# SetOptionBaseReactLoc

## Syntax

SapObject.SapModel.Results.Setup.SetOptionBaseReactLoc

## VB6 Procedure

Function SetOptionBaseReactLoc(ByVal gx As Double, ByVal gy As Double, ByVal gz As Double) As Long

## Parameters

gx, gy, gz

The global coordinates of the location at which the base reactions are reported.

## Remarks

This function sets the global coordinates of the location at which the base reactions are reported.

The function returns 0 if the coordinates are successfully set, otherwise it returns nonzero.

## VBA Example



```
Sub SetBaseReactLoc()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim gx As Double, gy As Double, gz As Double



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



   'set base reaction location

      ret = SapModel.Results.Setup.SetOptionBaseReactLoc(0, 0, 0)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetOptionBaseReactLoc](GetOptionBaseReactLoc.md)
