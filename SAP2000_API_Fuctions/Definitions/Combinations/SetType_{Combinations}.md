# SetTypeOAPI

## Syntax

SapObject.SapModel.RespCombo.SetTypeOAPI

## VB6 Procedure

Function SetTypeOAPI(ByVal Name As String, ByVal ComboType As Long) As Long

## Parameters

Name

The name of an existing load combination.

ComboType

This is 0, 1, 2, 3 or 4 indicating the load combination type.

0 = Linear Additive

1 = Envelope

2 = Absolute Additive

3 = SRSS

4 = Range Additive

## Remarks

This function sets the combination type for specified load combination.

The function returns zero if the type is successfully set, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetComboType()

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



   'add combo

      ret = SapModel.RespCombo.Add("COMB1", 0)



   'set combo type to Envelope

      ret = SapModel.RespCombo.SetTypeOAPI("COMB1", 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

Changed function name to SetTypeOAPI in v17.0.0.

## See Also

[GetTypeOAPI](SetType_%7BCombinations%7D.md)
