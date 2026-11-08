# GetDamperFrictionSpring

## Syntax

SapObject.SapModel.PropLink.GetDamperFrictionSpring

## VB6 Procedure

Function GetDamperFrictionSpring(ByVal Name As String, ByRef DOF() As Boolean, ByRef Fixed() As Boolean, ByRef NonLinear() As Boolean, ByRef Ke() As Double, ByRef Ce() As Double, ByRef k() As Double, ByRef k1() As Double, ByRef k2() As Double, ByRef u0() As Double, ByRef us() As Double, ByRef dir() As Long, ByRef dj2 As Double, ByRef dj3 As Double, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing friction spring damper-type link property.

DOF

This is a boolean array, dimensioned to 5, indicating if properties exist for a specified degree of freedom.

DOF(0) = U1

DOF(1) = U2

DOF(2) = U3

DOF(3) = R1

DOF(4) = R2

DOF(5) = R3

Fixed

This is a boolean array, dimensioned to 5, indicating if the specified degree of freedom is fixed (restrained.)

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

This is an array of initial (nonslipping) stiffness terms for the link property. The initial stiffness applies for nonlinear analyses.

k(0) = U1 [F/L]

k(1) = U2 [F/L]

k(2) = U3 [F/L]

k(3) = R1 [FL]

k(4) = R2 [FL]

k(5) = R3 [FL]

The term k(n) only applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

k1

This is an array of slipping stiffness when loading terms for the link property. The slipping stiffness when loading applies for nonlinear analyses.

k1(0) = U1 [F/L]

k1(1) = U2 [F/L]

k1(2) = U3 [F/L]

k1(3) = R1 [FL]

k1(4) = R2 [FL]

k1(5) = R3 [FL]

The term k1(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

k2

This is an array of slipping stiffness when unloading terms for the link property. The slipping stiffness when unloading applies for nonlinear analyses.

k2(0) = U1 [F/L]

k2(1) = U2 [F/L]

k2(2) = U3 [F/L]

k2(3) = R1 [FL]

k2(4) = R2 [FL]

k2(5) = R3 [FL]

The term k2(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

u0

This is an array of precompression displacement terms for the link property. The nonlinear precompression displacement applies for nonlinear analyses.

u0(0) = U1 [L]

u0(1) = U2 [L]

u0(2) = U3 [L]

u0(3) = R1

u0(4) = R2

u0(5) = R3

The term u0(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

us

This is an array of stop displacement terms for the link property. The nonlinear stop displacement applies for nonlinear analyses.

us(0) = U1 [L]

us(1) = U2 [L]

us(2) = U3 [L]

us(3) = R1

us(4) = R2

us(5) = R3

The term us(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

dj2

The distance from the J-End of the link to the U2 shear spring. This item applies only when DOF(1) = True. [L]

dj3

The distance from the J-End of the link to the U3 shear spring. This item applies only when DOF(2) = True. [L]

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property. If this item is input as Default, the program assigns a GUID to the property.

## **Remarks**

This function retrieves link property data for a friction spring damper-type link property.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkPropDamperFrictionSpring()
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
Dim MyK1() As Double
Dim MyK2() As Double
Dim MyU0() As Double
Dim MyUs() As Double
Dim MyDir() As Long
Dim DOF() As Boolean
Dim Fixed() As Boolean
Dim NonLinear() As Boolean
Dim Ke() As Double
Dim Ce() As Double
Dim K() As Double
Dim K1() As Double
Dim K2() As Double
Dim U0() As Double
Dim Us() As Double
Dim Dir() As Long
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
ReDim MyK1(5)
ReDim MyK2(5)
ReDim MyU0(5)
ReDim MyUs(5)
ReDim MyDir(5)
 
MyDOF(0) = True
MyKe(0) = 12
MyCe(0) = 0.01
 
MyDOF(1) = True
MyNonLinear(1) = True
MyKe(1) = 12
MyCe(1) = 0.01
MyK(1) = 20
MyK1(1)=2
MyK2(1) = 1
MyU0(1) = -.2
MyUs(1) = 1
MyDir(1) = 2
 
MyDOF(2) = True
MyFixed(2) = True
 
ret = SapModel.PropLink.SetDamperFrictionSpring("D1", MyDOF, MyFixed, MyNonLinear, MyKe, MyCe, MyK, MyK1, MyK2, MyU0, MyUs, MyDir, 1, 0)
 
'get link property data
ret = SapModel.PropLink.GetDamperFrictionSpring ("D1", DOF, Fixed, NonLinear, Ke, Ce, k, K1, K2, U0, Us, Dir, dj2, dj3, Notes, GUID)
 
'close Sap2000
SapObject.ApplicationExit False
Set SapModel = Nothing
Set SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 16.0.2.

## See Also

[SetDamperFrictionSpring](SetDamperFrictionSpring.md)
