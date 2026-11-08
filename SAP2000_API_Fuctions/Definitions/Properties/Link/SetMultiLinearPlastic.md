# SetMultiLinearPlastic

## Syntax

SapObject.SapModel.PropLink.SetMultiLinearPlastic

## VB6 Procedure

Function SetMultiLinearPlastic(ByVal Name As String, ByRef DOF() As Boolean, ByRef Fixed() As Boolean, ByRef NonLinear() As Boolean, ByRef Ke() As Double, ByRef Ce() As Double, ByVal dj2 As Double, ByVal dj3 As Double, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

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

dj2

The distance from the J-End of the link to the U2 shear spring. This item applies only when DOF(1) = True. [L]

dj3

The distance from the J-End of the link to the U3 shear spring. This item applies only when DOF(2) = True. [L]

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property. If this item is input as Default, the program assigns a GUID to the property.

## Remarks

This function initializes a multilinear plastic-type link property. If this function is called for an existing link property, all items for the property are reset to their default values.

The function returns zero if the property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetLinkPropMultiLinearPlastic()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyDOF() As Boolean

      Dim MyFixed() As Boolean

      Dim MyNonLinear() As Boolean

      Dim MyKe() As Double

      Dim MyCe() As Double



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



      MyDOF(0) = True

      MyKe(0) = 12

      MyCe(0) = 0.01



      MyDOF(1) = True

      MyNonLinear(1) = True

      MyKe(1) = 12

      MyCe(1) = 0.01



      MyDOF(2) = True

      MyFixed(2) = True



      ret = SapModel.PropLink.SetMultiLinearPlastic("MLP1", MyDOF, MyFixed, MyNonLinear, MyKe, MyCe, 2, 0)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetMultiLinearPlastic](GetMultiLinearPlastic.md)

[GetMultiLinearPoints](GetMultiLinearPoints.md)

[SetMultiLinearPoints](SetMultiLinearPoints.md)
