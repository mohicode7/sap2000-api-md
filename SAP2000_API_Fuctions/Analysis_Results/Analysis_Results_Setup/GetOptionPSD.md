# GetOptionPSD

## Syntax

SapObject.SapModel.Results.Setup.GetOptionPSD

## VB6 Procedure

Function GetOptionPSD(ByRef Value As Long) As Long

## Parameters

Value

This item is either 1 or 2

1 = RMS

2 = sqrt(PSD)

## Remarks

This function retrieves the output option for power spectral density results.

The function returns 0 if the output option is successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetPSDOutputOption()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value As Long



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



   'get output option

      ret = SapModel.Results.Setup.GetOptionPSD(Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetOptionPSD](SetOptionPSD.md)
