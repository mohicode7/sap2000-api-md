# GetPlasticWen

## Syntax

SapObject.SapModel.PropLink.GetPlasticWen

## VB6 Procedure

Function GetPlasticWen(ByVal Name As String, ByRef DOF() As Boolean, ByRef Fixed() As Boolean, ByRef NonLinear() As Boolean, ByRef Ke() As Double, ByRef Ce() As Double, ByRef k() As Double, ByRef Yield() As Double, ByRef Ratio() As Double, ByRef exp() As Double, ByRef dj2 As Double, ByRef dj3 As Double, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing plastic Wen-type link property.

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

NonLinear(3) = R1 has nonlinear properties

NonLinear(4) = R2 has nonlinear properties

NonLinear(5) = R3 has nonlinear properties

The term NonLinear(n) applies only when DOF(n) = True and Fixed(n) = False.

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

k(3) = R1 [FL]

k(4) = R2 [FL]

k(5) = R3 [FL]

The term k(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Yield

This is an array of yield force terms for the link property. The yield force applies for nonlinear analyses.

Yield(0) = U1 [F]

Yield(1) = U2 [F]

Yield(2) = U3 [F]

Yield(3) = R1 [FL]

Yield(4) = R2 [FL]

Yield(5) = R3 [FL]

The term Yield(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Ratio

This is an array of post-yield stiffness ratio terms for the link property. The post-yield stiffness ratio applies for nonlinear analyses. It is the post-yield stiffness divided by the initial stiffness.

Ratio(0) = U1

Ratio(1) = U2

Ratio(2) = U3

Ratio(3) = R1

Ratio(4) = R2

Ratio(5) = R3

The term Ratio(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

exp

This is an array of yield exponent terms for the link property. The yield exponent applies for nonlinear analyses. The yielding exponent that controls the sharpness of the transition from the initial stiffness to the yielded stiffness.

exp(0) = U1

exp(1) = U2

exp(2) = U3

exp(3) = R1

exp(4) = R2

exp(5) = R3

The term exp(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

dj2

The distance from the J-End of the link to the U2 shear spring. This item applies only when DOF(1) = True. [L]

dj3

The distance from the J-End of the link to the U3 shear spring. This item applies only when DOF(2) = True. [L]

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property.

## Remarks

This function retrieves link property data for a plastic Wen-type link property.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkPropPlasticWen()

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

      Dim MyYield() As Double

      Dim MyRatio() As Double

      Dim MyExp() As Double

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Nonlinear() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double

      Dim k() As Double

      Dim Yield() As Double

      Dim Ratio() As Double

      Dim exp() As Double

      Dim dj2 As Double

      Dim dj3 As Double

      Dim Notes As String

      Dim GUID As String



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

      ReDim MyYield(5)

      ReDim MyRatio(5)

      ReDim MyExp(5)



      MyDOF(0) = True

      MyKe(0) = 12

      MyCe(0) = 0.01



      MyDOF(1) = True

      MyNonLinear(1) = True

      MyKe(1) = 12

      MyCe(1) = 0.01

      MyK(1) = 20

      MyYield(1)= 50

      MyRatio(1)= 0.1

      MyExp(1)= 3



      MyDOF(2) = True

      MyFixed(2) = True



      ret = SapModel.PropLink.SetPlasticWen("PW1", MyDOF, MyFixed, MyNonLinear, MyKe, MyCe, MyK, MyYield, MyRatio, MyExp, 2, 0)



   'get link property data

      ret = SapModel.PropLink.GetPlasticWen("PW1", DOF, Fixed, NonLinear, Ke, Ce, k, Yield, Ratio, exp, dj2, dj3, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetPlasticWen](SetPlasticWen.md)
