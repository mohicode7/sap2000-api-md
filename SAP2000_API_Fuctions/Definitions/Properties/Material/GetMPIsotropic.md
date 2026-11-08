# GetMPIsotropic

## Syntax

SapObject.SapModel.PropMaterial.GetMPIsotropic

## VB6 Procedure

Function GetMPIsotropic(ByVal Name As String, ByRef e As Double, ByRef u As Double, ByRef a As Double, ByRef g As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing material property.

e

The modulus of elasticity. [F/L2]

u

Poisson’s ratio.

a

The thermal coefficient. [1/T]

g

The shear modulus. For isotropic materials this value is program calculated from the modulus of elasticity and poisson’s ratio. [F/L2]

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the mechanical properties for a material with an isotropic directional symmetry type.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value. The function returns an error if the symmetry type of the specified material is not isotropic.

## VBA Example



```
Sub GetMatPropIsotropic()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim e As Double

      Dim u As Double

      Dim a As Double

      Dim g As Double



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



   'get isotropic mechanical properties

      ret = SapModel.PropMaterial.GetMPIsotropic("Steel", e, u, a, g)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetMPIsotropic](SetMPIsotropic.md)
