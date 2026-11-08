# SetDamping

## Syntax

SapObject.SapModel.PropMaterial.SetDamping

## VB6 Procedure

Function SetDamping(ByVal Name As String, ByVal ModalRatio As Double, ByVal ViscousMassCoeff As Double, ByVal ViscousStiffCoeff As Double, ByVal HystereticMassCoeff As Double, ByVal HystereticStiffCoeff As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing material property.

ModalRatio

The modal damping ratio.

ViscousMassCoeff

The mass coefficient for viscous proportional damping.

ViscousStiffCoeff

The stiffness coefficient for viscous proportional damping.

HystereticMassCoeff

The mass coefficient for hysteretic proportional damping.

HystereticStiffCoeff

The stiffness coefficient for hysteretic proportional damping.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been previously defined for the material.

## Remarks

This function sets the additional material damping data for the material.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropDamping()

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

      ret = SapModel.PropMaterial.SetMaterial("Conc", MATERIAL_CONCRETE)



   'assign material damping data

      ret = SapModel.PropMaterial.SetDamping("Conc", 0.04, 0, 0, 0, 0)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetDamping](GetDamping.md)
