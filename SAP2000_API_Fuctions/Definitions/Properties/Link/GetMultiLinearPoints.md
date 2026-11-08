# GetMultiLinearPoints

## Syntax

SapObject.SapModel.PropLink.GetMultiLinearPoints

## VB6 Procedure

Function GetMultiLinearPoints(ByVal Name As String, ByVal DOF As Long, ByRef NumberPoints As Long, ByRef F() As Double, ByRef D() As Double, ByRef MyType As Long, ByRef a1 As Double, ByRef a2 As Double, ByRef b1 As Double, ByRef b2 As Double, ByRef eta As Double) As Long

## Parameters

Name

The name of an existing multilinear elastic or multilinear plastic link property.

DOF

This is 1, 2, 3, 4, 5 or 6, indicating the degree of freedom to which the multilinear points apply.

1 = U1

2 = U2

3 = U3

4 = R1

5 = R2

6 = R3

NumberPoints

The number of foce-defomation points for the specified degree of freedom.

F

This is an array, dimensioned to NumberPoints - 1, that includes the force at each point. When DOF is U1, U2 or U3, this is a force. When DOF is R1, R2 or R3, this is a moment. [F] if DOF <= 3, and [FL} if DOF > 3

D

This is an array, dimensioned to NumberPoints - 1, that includes the displacement at each point. When DOF is U1, U2 or U3, this is a translation. When DOF is R1, R2 or R3, this is a rotation. [L] if DOF <= 3, and [rad] if DOF > 3

MyType

This item applies only to multilinear plastic link properties. It is 1, 2 or 3, indicating the hysteresis type.

1 = Kinematic

2 = Takeda

3 = Pivot

a1

This item only applies to multilinear plastic link properties that have a pivot hysteresis type (MyType = 3). It is the Alpha1 hysteresis parameter.

a2

This item applies only to multilinear plastic link properties that have a pivot hysteresis type (MyType = 3). It is the Alpha2 hysteresis parameter.

b1

This item applies only to multilinear plastic link properties that have a pivot hysteresis type (MyType = 3). It is the Beta1 hysteresis parameter.

b2

This item applies only to multilinear plastic link properties that have a pivot hysteresis type (MyType = 3). It is the Beta2 hysteresis parameter.

eta

This item applies only to multilinear plastic link properties that have a pivot hysteresis type (MyType = 3). It is the Eta hysteresis parameter.

## Remarks

This function retrieves the force-deformation data for a specified degree of freedom in multilinear elastic and multilinear plastic link properties.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

To successfully retrieve this data from the indicated link property, the following conditions must be met:

1.    The link property must be multilinear elastic or multilinear plastic.

2.    The specified DOF must be active.

3.    The specified DOF must not be fixed.

4.    The specified DOF must be nonlinear.

## VBA Example



```
Sub GetLinkPropMultiLinearPoints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyDOF() As Boolean

      Dim MyFixed() As Boolean

      Dim MyNonLinear() As Boolean

      Dim MyKe() As Double

      Dim MyCe() As Double

      Dim MyF() As Double

      Dim MyD() As Double

      Dim NumberPoints As Long

      Dim F() As Double

      Dim D() As Double

      Dim MyType As Long

      Dim a1 As Double

      Dim a2 As Double

      Dim b1 As Double

      Dim b2 As Double

      Dim eta As Double



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



   'set multilinear force-defomation data

      ReDim MyF(4)

      ReDim MyD(4)



      MyF(0) = -12

      MyF(1) = -10

      MyF(2) = 0

      MyF(3) = 8

      MyF(4) = 9



      MyD(0) = -8

      MyD(1) = -0.6

      MyD(2) = 0

      MyD(3) = 0.2

      MyD(4) = 6



      ret = SapModel.PropLink.SetMultiLinearPoints("MLP1", 2, 5, MyF, MyD, 3, 9, 12, 0.75, 0.8, .1)



   'get multilinear force-defomation data

      ret = SapModel.PropLink.GetMultiLinearPoints("MLP1", 2, NumberPoints, F, D, MyType, a1, a2, b1, b2, eta)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetMultiLinearPoints](SetMultiLinearPoints.md)

[SetMultiLinearElastic](SetMultiLinearElastic.md)

[GetMultiLinearElastic](GetMultiLinearElastic.md)

[SetMultiLinearPlastic](SetMultiLinearPlastic.md)

[GetMultiLinearPlastic](GetMultiLinearPlastic.md)
