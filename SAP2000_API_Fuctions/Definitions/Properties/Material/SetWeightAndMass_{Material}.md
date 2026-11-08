# SetWeightAndMass

## Syntax

SapObject.SapModel.PropMaterial.SetWeightAndMass

## VB6 Procedure

Function SetWeightAndMass(ByVal Name As String, ByVal MyOption As Long, ByVal Value As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing material property.

MyOption

This is either 1 or 2, indicating what is specified by the Value item.

1 = Weight per unit volume is specified

2 = Mass per unit volume is specified

If the weight is specified, the corresponding mass is program calculated based on the specified weight. Similarly, if the mass is specified, the corresponding weight is program calculated based on the specified mass.

Value

This is either the weight per unit volume or the mass per unit volume, depending on the value of the MyOption item. [F/L3] for MyOption = 1 (weight), and [M/L3] for MyOption = 2 (mass)

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been define previously for the material.

## Remarks

This function assigns weight per unit volume or mass per unit volume to a material property.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropWeight()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'initialize new material property

      ret = SapModel.PropMaterial.SetMaterial("Steel", MATERIAL_STEEL)



   'assign material property weight per unit volume

      ret = SapModel.PropMaterial.SetWeightAndMass("Steel", 1, 0.00029)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetWeightAndMass](GetWeightAndMass_%7BMaterial%7D.md)
