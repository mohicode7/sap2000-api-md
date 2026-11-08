# GetTemp

## Syntax

SapObject.SapModel.PropMaterial.GetTemp

## VB6 Procedure

Function GetTemp(ByVal Name As String, ByRef NumberItems as long, ByRef Temp() As Double) As Long

## Parameters

Name

The name of a material property.

NumberItems

The number of different temperatures at which properties are specified for the material.

Temp

This is an array that includes the different temperatures at which properties are specified for the material.

## Remarks

This function retrieves the temperatures at which properties are specified for a material.

The function returns zero if the temperatures are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetMatPropTemps()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim MyTemp() As Double

      Dim Temp() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'initialize new material property

      ret = SapModel.PropMaterial.SetMaterial("Steel", MATERIAL_STEEL)



   'specify temps at which properties will be provided

      ReDim MyTemp(2)

      MyTemp(0) = 0

      MyTemp(1) = 50

      MyTemp(2) = 100

      ret = SapModel.PropMaterial.SetTemp("Steel", 3, MyTemp)



   'get temps at which properties are provided

      ret = SapModel.PropMaterial.GetTemp("Steel", NumberItems, Temp)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetTemp](SetTemp.md)
