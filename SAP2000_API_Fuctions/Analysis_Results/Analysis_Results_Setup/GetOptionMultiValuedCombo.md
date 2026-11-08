# GetOptionMultiValuedCombo

## Syntax

SapObject.SapModel.Results.Setup.GetOptionMultiValuedCombo

## VB6 Procedure

Function GetOptionMultiValuedCombo(ByRef Value As Long) As Long

## Parameters

Value

This item is either 1, 2, or 3

1 = Envelopes

2 = Multiple values, if possible

3 = Correspondence

## Remarks

This function retrieves the output option for multi-valued load combination results.

The function returns 0 if the output option is successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetMultiValuedComboOutputOption()

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

      ret = SapModel.Results.Setup.GetOptionMultiValuedCombo(Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

Added option for correspondence in version 16.02.

## See Also

[SetOptionMultiValuedCombo](SetOptionMultiValuedCombo.md)
