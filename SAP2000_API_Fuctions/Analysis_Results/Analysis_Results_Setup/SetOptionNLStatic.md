# SetOptionNLStatic

## Syntax

SapObject.SapModel.Results.Setup.SetOptionNLStatic

## VB6 Procedure

Function SetOptionNLStatic(ByVal Value As Long) As Long

## Parameters

Value

This item is 1, 2 or 3

1 = Envelopes

2 = Step-by-Step

3 = Last Step

## Remarks

This function sets the output option for nonlinear static results.

The function returns 0 if the output option is successfully set, otherwise it returns nonzero.

## VBA Example



```
Sub SetNonlinearStaticOutputOption()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'set output option

      ret = SapModel.Results.Setup.SetOptionNLStatic(1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetOptionNLStatic](GetOptionNLStatic.md)
