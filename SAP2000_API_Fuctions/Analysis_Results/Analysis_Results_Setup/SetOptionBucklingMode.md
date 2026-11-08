# SetOptionBucklingMode

## Syntax

SapObject.SapModel.Results.Setup.SetOptionBucklingMode

## VB6 Procedure

Function SetOptionBucklingMode(ByVal BuckModeStart As Long, ByVal BuckModeEnd As Long, Optional ByVal BuckModeAll As Boolean = False) As Long

## Parameters

BuckModeStart

The first buckling mode for which the buckling factor is reported when the BuckModeAll item is False.

BuckModeEnd

The last buckling mode for which the buckling factor is reported when the BuckModeAll item is False.

BuckModeAll

If this item is True, buckling factors are reported for all calculated buckling modes. If it is False, buckling factors are reported for the buckling modes indicated by the BuckModeStart and BuckModeEnd items.

## Remarks

This function sets the buckling modes for which buckling factors are reported.

The function returns 0 if the modes are successfully set, otherwise it returns nonzero.

## VBA Example



```
Sub SetBucklingModeData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim BuckModeStart As Long

      Dim BuckModeEnd As Long

      Dim BuckModeAll As Boolean



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



   'set buckling mode data

      ret = SapModel.Results.Setup.SetOptionBucklingMode(1, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetOptionBucklingMode](GetOptionBucklingMode.md)
