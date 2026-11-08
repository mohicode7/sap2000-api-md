# GetWeightAndMass

## Syntax

SapObject.SapModel.PropMaterial.GetWeightAndMass

## VB6 Procedure

Function GetWeightAndMass(ByVal Name As String, ByRef w As Double, ByRef m As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing material property.

w

The weight per unit volume for the material. [F/L3]

m

The mass per unit volume for the material. [M/L3]

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the weight per unit volume and mass per unit volume of the material.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetMatPropWeightAndMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim w As Double

      Dim m As Double



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



   'get material weight and mass per unit volume

      ret = SapModel.PropMaterial.GetWeightAndMass("Steel", w, m)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetWeightAndMass](SetWeightAndMass_%7BMaterial%7D.md)
