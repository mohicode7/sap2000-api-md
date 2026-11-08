# GetLinear

## Syntax

SapObject.SapModel.PropLink.GetLinear

## VB6 Procedure

Function GetLinear(ByVal Name As String, ByRef DOF() As Boolean, ByRef Fixed() As Boolean, ByRef Ke() As Double, ByRef Ce() As Double, ByRef dj2 As Double, ByRef dj3 As Double, ByRef KeCoupled As Boolean, ByRef CeCoupled As Boolean, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing linear-type link property.

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

Fixed(0) = U1 fixity if DOF(0) = True

Fixed(1) = U2 fixity if DOF(1) = True

Fixed(2) = U3 fixity if DOF(2) = True

Fixed(3) = R1 fixity if DOF(3) = True

Fixed(4) = R2 fixity if DOF(4) = True

Fixed(5) = R3 fixity if DOF(5) = True

Ke

This is an array of stiffness terms for the link property. There are 6 terms in the array if the stiffness is uncoupled and 21 if it is coupled. The KeCoupled item indicates if the stiffness is coupled.

If the stiffness is uncoupled:

Ke(0) = U1 [F/L]

Ke(1) = U2 [F/L]

Ke(2) = U3 [F/L]

Ke(3) = R1 [FL]

Ke(4) = R2 [FL]

Ke(5) = R3 [FL]

If the stiffness is coupled:

Ke(0) = U1U1 [F/L]

Ke(1) = U1U2 [F/L]

Ke(2) = U2U2 [F/L]

Ke(3) = U1U3 [F/L]

Ke(4) = U2U3 [F/L]

Ke(5) = U3U3 [F/L]

Ke(6) = U1R1 [F]

Ke(7) = U2R1 [F]

Ke(8) = U3R1 [F]

Ke(9) = R1R1 [FL]

Ke(10) = U1R2 [F]

Ke(11) = U2R2 [F]

Ke(12) = U3R2 [F]

Ke(13) = R1R2 [FL]

Ke(14) = R2R2 [FL]

Ke(15) = U1R3 [F]

Ke(16) = U2R3 [F]

Ke(17) = U3R3 [F]

Ke(18) = R1R3 [FL]

Ke(19) = R2R3 [FL]

Ke(20) = R3R3 [FL]

Ce

This is an array of damping terms for the link property. There are 6 terms in the array if the damping is uncoupled and 21 if it is coupled. The CeCoupled item indicates if the damping is coupled.

If the damping is uncoupled:

Ce(0) = U1 [F/L]

Ce(1) = U2 [F/L]

Ce(2) = U3 [F/L]

Ce(3) = R1 [FL]

Ce(4) = R2 [FL]

Ce(5) = R3 [FL]

If the damping is coupled:

Ce(0) = U1U1 [F/L]

Ce(1) = U1U2 [F/L]

Ce(2) = U2U2 [F/L]

Ce(3) = U1U3 [F/L]

Ce(4) = U2U3 [F/L]

Ce(5) = U3U3 [F/L]

Ce(6) = U1R1 [F]

Ce(7) = U2R1 [F]

Ce(8) = U3R1 [F]

Ce(9) = R1R1 [FL]

Ce(10) = U1R2 [F]

Ce(11) = U2R2 [F]

Ce(12) = U3R2 [F]

Ce(13) = R1R2 [FL]

Ce(14) = R2R2 [FL]

Ce(15) = U1R3 [F]

Ce(16) = U2R3 [F]

Ce(17) = U3R3 [F]

Ce(18) = R1R3 [FL]

Ce(19) = R2R3 [FL]

Ce(20) = R3R3 [FL]

dj2

The distance from the J-End of the link to the U2 shear spring. This item applies only when DOF(1) = True. [L]

dj3

The distance from the J-End of the link to the U3 shear spring. This item applies only when DOF(2) = True. [L]

KeCoupled

This item is True if the link stiffness, Ke, is coupled. There are 21 terms in the Ke array if Ke is coupled; otherwise there are 6 terms.

CeCoupled

This item is True if the link damping, Ce, is coupled. There are 21 terms in the Ce array if Ce is coupled; otherwise there are 6 terms.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property.

## Remarks

This function retrieves link property data for a linear-type link property.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkPropLinear()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyDOF() As Boolean

      Dim MyFixed() As Boolean

      Dim MyKe() As Double

      Dim MyCe() As Double

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double

      Dim dj2 As Double

      Dim dj3 As Double

      Dim KeCoupled As Boolean

      Dim CeCoupled As Boolean

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

      ReDim MyKe(5)

      ReDim MyCe(5)

      MyDOF(0) = True

      MyKe(0) = 12

      ret = SapModel.PropLink.SetLinear("L1", MyDOF, MyFixed, MyKe, MyCe, 0, 0)



   'get link property data

      ret = SapModel.PropLink.GetLinear("L1", DOF, Fixed, Ke, Ce, dj2, dj3, KeCoupled, CeCoupled, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetLinear](SetLinear.md)
