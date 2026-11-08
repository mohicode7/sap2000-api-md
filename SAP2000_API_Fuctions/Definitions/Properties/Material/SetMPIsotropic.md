# SetMPIsotropic

## Syntax

SapObject.SapModel.PropMaterial.SetMPIsotropic

## VB6 Procedure

Function SetMPIsotropic(ByVal Name As String, ByVal e As Double, ByVal u As Double, ByVal a As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing material property.

e

The modulus of elasticity. [F/L2]

u

Poisson’s ratio.

a

The thermal coefficient. [1/T]

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.

## Remarks

This function sets the material directional symmetry type to isotropic, and assigns the isotropic mechanical properties.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropIsotropic()

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



   'assign isotropic mechanical properties

      ret = SapModel.PropMaterial.SetMPIsotropic("Steel", 29500, 0.25, 6E-06)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetMPIsotropic](GetMPIsotropic.md)
