# GetTriplePendulumIsolator

## Syntax

SapObject.SapModel.PropLink.GetTriplePendulumIsolator

## VB6 Procedure

Function GetTriplePendulumIsolator(ByVal Name As String, ByRef dof() As Boolean, ByRef Fixed() As Boolean, ByRef Nonlinear() As Boolean, ByRef Ke() As Double, ByRef Ce() As Double, ByRef K1 As Double, ByRef Damping As Double, ByRef K() As Double, ByRef Slow() As Double, ByRef Fast() As Double, ByRef Rate() As Double, ByRef Radius() As Double, ByRef StopDist() As Double, ByRef HeightOut As Double, ByRef HeightIn As Double, ByRef dj2 As Double, ByRef dj3 As Double, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing Triple Pendulum Isolator type link property.

DOF

This is a Boolean array, dimensioned to 5, indicating if properties exist for a specified degree of freedom.

DOF(0) = U1

DOF(1) = U2

DOF(2) = U3

DOF(3) = R1

DOF(4) = R2

DOF(5) = R3

Fixed

This is a Boolean array, dimensioned to 5, indicating if the specified degree of freedom is fixed (restrained).

Fixed(0) = U1 fixity

Fixed(1) = U2 fixity

Fixed(2) = U3 fixity

Fixed(3) = R1 fixity

Fixed(4) = R2 fixity

Fixed(5) = R3 fixity

The term Fixed(n) only applies when DOF(n) = True.

NonLinear

This is a Boolean array, dimensioned to 5, indicating if nonlinear properties exist for a specified degree of freedom.

NonLinear(0) = U1 has nonlinear properties

NonLinear(1) = U2 has nonlinear properties

NonLinear(2) = U3 has nonlinear properties

NonLinear(3) = R1, Not Used

NonLinear(4) = R2, Not Used

NonLinear(5) = R3, Not Used

Note that this item is applicable for degrees of freedom U1, U2 and U3 only. For those degrees of freedom, the term NonLinear(n) applies only when DOF(n) = True and Fixed(n) = False.

Ke

This is an array of effective stiffness terms for the link property. The effective stiffness applies for linear analyses, and also for nonlinear analysis for those DOF for which NonLinear(n) = False.

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

K1

This is the axial compression stiffness for the U1 degree of freedom. This item applies for nonlinear analyses. [F/L]

Damping

This is the nonlinear damping coefficient for the axial degree of freedom, U1, when it is in compression. This item applies for nonlinear analyses. [F/L]

K

This is an array, dimensioned to 3, of initial nonlinear stiffness (before sliding) for each sliding surface.

K(0) = for the outer top sliding surface [F/L]

K(1) = for the outer bottom sliding surface [F/L]

K(2) = for the inner top sliding surface [F/L]

K(3) = for the inner bottom sliding surface [F/L]

Note that this item is applicable for degrees of freedom U2 and U3 only. For those degrees of freedom, the term k(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Slow

This is an array, dimensioned to 3, of the friction coefficient at zero velocity for each sliding surface when U1 is in compression.

Slow(0) = for the outer top sliding surface

Slow(1) = for the outer bottom sliding surface

Slow(2) = for the inner top sliding surface

Slow(3) = for the inner bottom sliding surface

Note that this item is applicable for degrees of freedom U2 and U3 only. For those degrees of freedom, the term Slow(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Fast

This is an array, dimensioned to 3, of the friction coefficient at fast velocity for each sliding surface when U1 is in compression.

Fast(0) = for the outer top sliding surface

Fast(1) = for the outer bottom sliding surface

Fast(2) = for the inner top sliding surface

Fast(3) = for the inner bottom sliding surface

Note that this item is applicable for degrees of freedom U2 and U3 only. For those degrees of freedom, the term Fast(n) applies only  when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Rate

This is an array, dimensioned to 3, of the inverse of the characteristic sliding velocity for the Slow and Fast friction coefficients for each sliding surface. This item applies for nonlinear analyses.

Rate(0) = for the outer top sliding surface [s/L]

Rate(1) = for the outer bottom sliding surface [s/L]

Rate(2) = for the inner top sliding surface [s/L]

Rate(3) = for the inner bottom sliding surface [s/L]

Note that this item is applicable for degrees of freedom U2 and U3 only. For those degrees of freedom, the term Rate(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Radius

This is an array, dimensioned to 3, of the radius for each sliding surface. Inputting 0 means there is an infinite radius, that is, the slider is flat. This item applies for nonlinear analyses.

Radius(0) = for the outer top sliding surface [L]

Radius(1) = for the outer bottom sliding surface [L]

Radius(2) = for the inner top sliding surface [L]

Radius(3) = for the inner bottom sliding surface [L]

Note that this item is applicable for degrees of freedom U2 and U3 only. For those degrees of freedom, the term Radius(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

StopDist

This is an array, dimensioned to 3, of the amount of displacement allowed before hitting a stiff limit for each sliding surface. Inputting 0 means there is no stop. This item applies for nonlinear analyses.

StopDist(0) = for the outer top sliding surface [L]

StopDist(1) = for the outer bottom sliding surface [L]

StopDist(2) = for the inner top sliding surface [L]

StopDist(3) = for the inner bottom sliding surface [L]

Note that this item is applicable for degrees of freedom U2 and U3 only. For those degrees of freedom, the term StopDist(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

HeightOut

This is the height (distance) between the outer sliding surfaces at zero displacement. [L]

Note that this item is applicable for degrees of freedom U2 and U3 only. For those degrees of freedom, the term applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

HeightIn

This is the height (distance) between the inner sliding surfaces. [L]

Note that this item is applicable for degrees of freedom U2 and U3 only. For those degrees of freedom, the term applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

dj2

The distance from the J-End of the link to the U2 shear spring, that is, the center of the isolator. This item applies only when DOF(2) = True. [L]

dj3

The distance from the J-End of the link to the U3 shear spring, that is, the center of the isolator. This item applies only when DOF(3) = True. [L]

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property.

## Remarks

This function retrieves link property data for a Triple Pendulum Isolator type link property.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkPropTriplePendulumIsolator()

   'dimension variables

      Dim SapObject As Sap2000.SapObject

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

      Dim MyStopDist() As Double

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Nonlinear() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double

      Dim K() As Double

      Dim Slow() As Double

      Dim Fast() As Double

      Dim Rate() As Double

      Dim Radius() As Double

      Dim StopDist() As Double

      Dim K1 As Double

      Dim Damping As Double

      Dim Hout As Double

      Dim Hin As Double

      Dim dj2 As Double

      Dim dj3 As Double

      Dim Notes As String

      Dim GUID As String



   'create Sap2000 object

      Set SapObject = New SAP2000.SapObject



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

      ReDim MyK(3)

      ReDim MySlow(3)

      ReDim MyFast(3)

      ReDim MyRate(3)

      ReDim MyRadius(3)

      ReDim MyStopDist(3)



      MyDOF(0) = True

      MyNonLinear(0) = True

      MyKe(0) = 12

      MyCe(0) = 0.01



      MyDOF(1) = True

      MyNonLinear(1) = True

      MyKe(1) = 12

      MyCe(1) = 0.01



      MyDOF(2) = True

      MyNonLinear(2) = True

      MyKe(2) = 14

      MyCe(2) = 0.008



      MyK(0) = 20

      MySlow(0)= 0.6

      MyFast(0)= 0.5

      MyRate(0)= 10

      MyRadius(0)= 80

      MyStopDist(0)= 8



      MyK(1) = 22

      MySlow(1)= 0.66

      MyFast(1)= 0.55

      MyRate(1)= 12

      MyRadius(1)= 75

      MyStopDist(1)= 8



      MyK(2) = 20

      MySlow(2)= 0.6

      MyFast(2)= 0.5

      MyRate(2)= 10

      MyRadius(2)= 80

      MyStopDist(2)= 8



      MyK(3) = 20

      MySlow(3)= 0.6

      MyFast(3)= 0.5

      MyRate(3)= 10

      MyRadius(3)= 80

      MyStopDist(3)= 8



      MyDOF(3) = True

      MyKe(3) = 15

      MyCe(3) = 0



      MyDOF(4) = True

      MyFixed(4) = True



      ret = SapModel.PropLink.SetTriplePendulumIsolator("TPI1", MyDOF, MyFixed, MyNonLinear, MyKe, MyCe, 1000, 0.1, MyK, MySlow, MyFast, MyRate, MyRadius, MyStopDist, 6, 3, 3, 3)



   'get link property data

      ret = SapModel.PropLink.GetTriplePendulumIsolator("TPI1", DOF, Fixed, NonLinear, Ke, Ce, K1, Damping, K, Slow, Fast, Rate, Radius, StopDist, Hout, Hin, dj2, dj3, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.00.

## See Also

[SetTriplePendulumIsolator](SetTriplePendulumIsolator.md)
