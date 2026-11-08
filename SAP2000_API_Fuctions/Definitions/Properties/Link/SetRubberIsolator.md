# SetRubberIsolator

## Syntax

SapObject.SapModel.PropLink.SetRubberIsolator

## VB6 Procedure

Function SetRubberIsolator(ByVal Name As String, ByRef DOF() As Boolean, ByRef Fixed() As Boolean, ByRef NonLinear() As Boolean, ByRef Ke() As Double, ByRef Ce() As Double, ByRef k() As Double, ByRef Yield() As Double, ByRef Ratio() As Double, ByVal dj2 As Double, ByVal dj3 As Double, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

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

NonLinear(0) = U1, Not Used

NonLinear(1) = U2 has nonlinear properties

NonLinear(2) = U3 has nonlinear properties

NonLinear(3) = R1, Not Used

NonLinear(4) = R2, Not Used

NonLinear(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U2 and U3. For those degrees of freedom, the term NonLinear(n) applies only when DOF(n) = True and Fixed(n) = False.

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

k(0) = U1, Not Used

k(1) = U2 [F/L]

k(2) = U3 [F/L]

k(3) = R1, Not Used

k(4) = R2, Not Used

k(5) = R3, Not Used

Note that this item is applicable only for degrees of freedom U2 and U3. For those degrees of freedom, the term k(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Yield

This is an array of yield force terms for the link property. The yield force applies for nonlinear analyses.

Yield(0) = U1, Not Used

Yield(1) = U2 [F]

Yield(2) = U3 [F]

Yield(3) = R1, Not Used

Yield(4) = R2, Not Used

Yield(5) = R3, Not Used

The term Yield(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

Ratio

This is an array of post-yield stiffness ratio terms for the link property. The post-yield stiffness ratio applies for nonlinear analyses. It is the post-yield stiffness divided by the initial stiffness.

Ratio(0) = U1, Not Used

Ratio(1) = U2

Ratio(2) = U3

Ratio(3) = R1, Not Used

Ratio(4) = R2, Not Used

Ratio(5) = R3, Not Used

The term Ratio(n) applies only when DOF(n) = True, Fixed(n) = False and NonLinear(n) = True.

dj2

The distance from the J-End of the link to the U2 shear spring. This item applies only when DOF(1) = True. [L]

dj3

The distance from the J-End of the link to the U3 shear spring. This item applies only when DOF(2) = True. [L]

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property. If this item is input as Default, the program assigns a GUID to the property.

## Remarks

This function initializes a rubber isolator-type link property. If this function is called for an existing link property, all items for the property are reset to their default value.

The function returns zero if the property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetLinkPropRubberIsolator()

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



      MyDOF(2) = True

      MyNonLinear(2) = True

      MyKe(2) = 15

      MyCe(2) = 0.008

      MyK(2) = 22

      MyYield(2)= 60

      MyRatio(2)= 0.15



      MyDOF(3) = True

      MyFixed(3) = True



      ret = SapModel.PropLink.SetRubberIsolator("RI1", MyDOF, MyFixed, MyNonLinear, MyKe, MyCe, MyK, MyYield, MyRatio, 2, 3)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetRubberIsolator](GetRubberIsolator.md)
