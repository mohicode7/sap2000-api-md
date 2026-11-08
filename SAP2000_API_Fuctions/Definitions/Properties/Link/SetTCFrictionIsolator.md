# SetTCFrictionIsolator

## Syntax

SapObject.SapModel.PropLink.SetTCFrictionIsolator

## VB6 Procedure

Function SetTCFrictionIsolator(ByVal Name As String, ByRef DOF() As Boolean, ByRef Fixed() As Boolean, ByRef NonLinear() As Boolean, ByRef Ke() As Double, ByRef Ce() As Double, ByRef k() As Double, ByRef Slow() As Double, ByRef Fast() As Double, ByRef Rate() As Double, ByRef Radius() As Double, ByRef SlowT() As Double, ByRef FastT() As Double, ByRef RateT() As Double, ByVal kt As Double, ByVal dis As Double, ByVal dist As Double, ByVal Damping As Double, ByVal dj2 As Double, ByVal dj3 As Double, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new link property. If this is an existing property, that property is modified; otherwise, a new property is added.

DOF

This is a boolean array, dimensioned to 5, indicating if properties exist for a specified degree of freedom.

DOF(0) = U1

DOF(1) = U2

DOF(2) = U3

DOF(3) = R1

DOF(4) = R2

DOF(5) = R3

Fixed

This is a boolean array, dimensioned to 5, indicating if the specified degree of freedom is fixed (restrained).

Fixed(0) = U1 fixity

Fixed(1) = U2 fixity

Fixed(2) = U3 fixity

Fixed(3) = R1 fixity

Fixed(4) = R2 fixity

Fixed(5) = R3 fixity

The term Fixed(n) applies only when DOF(n) = True.

NonLinear

This is a boolean array, dimensioned to 5, indicating if nonlinear properties exist for a specified degree of freedom.

NonLinear(0) = U1 has nonlinear properties

NonLinear(1) = U2 has nonlinear properties

NonLinear(2) = U3 has nonlinear properties

NonLinear(3) = R1, Not Used

NonLinear(4) = R2, Not Used

NonLinear(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U1, U2 and U3. For those degrees of freedom, the term NonLinear(n) applies only when DOF(n) = True and Fixed(n) = False.

Ke

This is an array of effective stiffness terms for the link property. The effective stiffness applies for linear analyses.

Ke(0) = U1 [F/L]

Ke(1) = U2 [F/L]

Ke(2) = U3 [F/L]

Ke(3) = R1 [FL]

Ke(4) = R2 [FL]

Ke(5) = R3 [FL]

The term Ke(n) applies only when DOF(n) = True and Fixed(n) = False.

Ce

This is an array of effective damping terms for the link property. The effective damping applies for linear analyses.

Ce(0) = U1 [F/L]

Ce(1) = U2 [F/L]

Ce(2) = U3 [F/L]

Ce(3) = R1 [FL]

Ce(4) = R2 [FL]

Ce(5) = R3 [FL]

The term Ce(n) applies only when DOF(n) = True and Fixed(n) = False.

k

This is an array of initial stiffness terms for the link property. The initial stiffness applies for nonlinear analyses.

k(0) = U1 [F/L]

k(1) = U2 [F/L]

k(2) = U3 [F/L]

k(3) = R1, Not Used

k(4) = R2, Not Used

k(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U1, U2 and U3. For those degrees of freedom, the term k(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Slow

This is an array of the friction coefficient at zero velocity terms when U1 is in compression for the link property. This coefficient applies for nonlinear analyses.

Slow(0) = U1, Not Used

Slow(1) = U2

Slow(2) = U3

Slow(3) = R1, Not Used

Slow(4) = R2, Not Used

Slow(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U2 and U3. For those degrees of freedom, the term Slow(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Fast

This is an array of the friction coefficient at fast velocity terms when U1 is in compression for the link property. This coefficient applies for nonlinear analyses.

Fast(0) = U1, Not Used

Fast(1) = U2

Fast(2) = U3

Fast(3) = R1, Not Used

Fast(4) = R2, Not Used

Fast(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U2 and U3. For those degrees of freedom, the term Fast(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Rate

This is an array of the inverse of the characteristic sliding velocity terms when U1 is in compression for the link property. This item applies for nonlinear analyses.

Rate(0) = U1, Not Used

Rate(1) = U2 [s/L]

Rate(2) = U3 [s/L]

Rate(3) = R1, Not Used

Rate(4) = R2, Not Used

Rate(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U2 and U3. For those degrees of freedom, the term Rate(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Radius

This is an array of the radius of the sliding contact surface terms for the link property. Inputting 0 means there is an infinite radius, that is, the slider is flat. This item applies for nonlinear analyses.

Radius(0) = U1, Not Used

Radius(1) = U2 [L]

Radius(2) = U3 [L]

Radius(3) = R1, Not Used

Radius(4) = R2, Not Used

Radius(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U2 and U3. For those degrees of freedom, the term Radius(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

SlowT

This is an array of the friction coefficient at zero velocity terms when U1 is in tension for the link property. This coefficient applies for nonlinear analyses.

SlowT(0) = U1, Not Used

SlowT(1) = U2

SlowT(2) = U3

SlowT(3) = R1, Not Used

SlowT(4) = R2, Not Used

SlowT(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U2 and U3. For those degrees of freedom, the term SlowT(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

FastT

This is an array of the friction coefficient at fast velocity terms when U1 is in tension for the link property. This coefficient applies for nonlinear analyses.

FastT(0) = U1, Not Used

FastT(1) = U2

FastT(2) = U3

FastT(3) = R1, Not Used

FastT(4) = R2, Not Used

FastT(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U2 and U3. For those degrees of freedom, the term FastT(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

RateT

This is an array of the inverse of the characteristic sliding velocity terms when U1 is in tension for the link property. This item applies for nonlinear analyses.

RateT(0) = U1, Not Used

RateT(1) = U2 [s/L]

RateT(2) = U3 [s/L]

RateT(3) = R1, Not Used

RateT(4) = R2, Not Used

RateT(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U2 and U3. For those degrees of freedom, the term RateT(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

kt

The axial translational tension stiffness for the U1 degree of freedom. This item applies for nonlinear analyses. [F/L]

dis

The U1 degree of freedom gap opening for compression. This item applies for nonlinear analyses. [L]

dist

The U1 degree of freedom gap opening for tension. This item applies for nonlinear analyses. [L]

Damping

The nonlinear damping coefficient used for the axial translational degree of freedom, U1. This item applies for nonlinear analyses. [F/L]

dj2

The distance from the J-End of the link to the U2 shear spring. This item applies only when DOF(1) = True. [L]

dj3

The distance from the J-End of the link to the U3 shear spring. This item applies only when DOF(2) = True. [L]

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property. If this item is input as Default, the program assigns a GUID to the property.

## Remarks

This function initializes a T/C friction isolator-type link property. If this function is called for an existing link property, all items for the property are reset to their default value.

The function returns zero if the property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetLinkPropTCFrictionIsolator()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyDOF() As Boolean

      Dim MyFixed() As Boolean

      Dim MyNonLinear() As Boolean

      Dim MyKe() As Double

      Dim MyCe() As Double

      Dim MyK() As Double

      Dim MySlow() As Double

      Dim MyFast() As Double

      Dim MyRate() As Double

      Dim MyRadius() As Double

      Dim MySlowT() As Double

      Dim MyFastT() As Double

      Dim MyRateT() As Double



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



   'add link property

      ReDim MyDOF(5)

      ReDim MyFixed(5)

      ReDim MyNonLinear(5)

      ReDim MyKe(5)

      ReDim MyCe(5)

      ReDim MyK(5)

      ReDim MySlow(5)

      ReDim MyFast(5)

      ReDim MyRate(5)

      ReDim MyRadius(5)

      ReDim MySlowT(5)

      ReDim MyFastT(5)

      ReDim MyRateT(5)



      MyDOF(0) = True

      MyNonLinear(0) = True

      MyKe(0) = 12

      MyCe(0) = 0.01

      MyK(0) = 1000



      MyDOF(1) = True

      MyNonLinear(1) = True

      MyKe(1) = 12

      MyCe(1) = 0.01

      MyK(1) = 20

      MySlow(1)= 0.6

      MyFast(1)= 0.5

      MyRate(1)= 10

      MyRadius(1)= 80

      MySlowT(1)= 0.61

      MyFastT(1)= 0.51

      MyRateT(1)= 10.1



      MyDOF(2) = True

      MyNonLinear(2) = True

      MyKe(2) = 14

      MyCe(2) = 0.008

      MyK(2) = 22

      MySlow(2)= 0.66

      MyFast(2)= 0.55

      MyRate(2)= 12

      MyRadius(2)= 75

      MySlowT(2)= 0.67

      MyFastT(2)= 0.56

      MyRateT(2)= 12.1



      MyDOF(3) = True

      MyKe(3) = 15

      MyCe(3) = 0



      MyDOF(4) = True

      MyFixed(4) = True



      ret = SapModel.PropLink.SetTCFrictionIsolator("TCFI1", MyDOF, MyFixed, MyNonLinear, MyKe, MyCe, MyK, MySlow, MyFast, MyRate, MyRadius, MySlowT, MyFastT, MyRateT, 18, 2, 3, 0.1, 2, 3)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetTCFrictionIsolator](GetTCFrictionIsolator.md)
