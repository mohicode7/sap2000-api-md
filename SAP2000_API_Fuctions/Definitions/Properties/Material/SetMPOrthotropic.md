# SetMPOrthotropic

## Syntax

SapObject.SapModel.PropMaterial.SetMPOrthotropic

## VB6 Procedure

Function SetMPOrthotropic(ByVal Name As String, ByRef e() As Double, ByRef u() As Double, ByRef a() As Double, ByRef g() As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing material property.

e

This is an array that includes the modulus of elasticity.

e(0) = E1 [F/L2]

e(1) = E2 [F/L2]

e(2) = E3 [F/L2]

u

This is an array that includes poisson’s ratio.

u(0) = U12

u(1) = U13

u(2) = U23

a

This is an array that includes the thermal coefficient.

a(0) = A1 [1/T]

a(1) = A2 [1/T]

a(2) = A3 [1/T]

g

This is an array that includes the shear modulus.

g(0) = G12 [F/L2]

g(1) = G13 [F/L2]

g(2) = G23 [F/L2]

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.

## Remarks

This function sets the material directional symmetry type to orthotropic, and assigns the orthotropic mechanical properties.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropOrthotropic()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyE() As Double

      Dim MyU() As Double

      Dim MyA() As Double

      Dim MyG() As Double



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



   'assign orthotropic mechanical properties

      ReDim MyE(2)

      ReDim MyU(2)

      ReDim MyA(2)

      ReDim MyG(2)

      MyE(0)=30000

      MyE(1)=10000

      MyE(2)=2000

      MyU(0)=0.2

      MyU(1)=0.05

      MyU(2)=0.1

      MyA(0)=6.5E-6

      MyA(1)=6.5E-6

      MyA(2)=6.5E-6

      MyG(0)=1500

      MyG(1)=2500

      MyG(2)=8700

      ret = SapModel.PropMaterial.SetMPOrthotropic("Steel", MyE, MyU, MyA, MyG)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetMPOrthotropic](GetMPOrthotropic.md)
