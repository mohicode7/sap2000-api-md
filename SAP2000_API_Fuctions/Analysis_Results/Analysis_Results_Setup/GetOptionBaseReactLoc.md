# GetOptionBaseReactLoc

## Syntax

SapObject.SapModel.Results.Setup.GetOptionBaseReactLoc

## VB6 Procedure

Function GetOptionBaseReactLoc(ByRef gx As Double, ByRef gy As Double, ByRef gz As Double) As Long

## Parameters

gx, gy, gz

The global coordinates of the location at which the base reactions are reported.

## Remarks

This function retrieves the global coordinates of the location at which the base reactions are reported.

The function returns 0 if the coordinates are successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetBaseReactLoc()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim gx As Double, gy As Double, gz As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create a SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'get base reaction location

      ret = SapModel.Results.Setup.GetOptionBaseReactLoc(gx, gy, gz)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetOptionBaseReactLoc](SetOptionBaseReactLoc.md)
