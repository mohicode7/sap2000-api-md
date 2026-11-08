# GetSpringCoupled

## Syntax

SapObject.SapModel.PointObj.GetSpringCoupled

## VB6 Procedure

Function GetSpringCoupled(ByVal Name As String, ByRef k() As Double) As Long

## Parameters

Name

The name of an existing point object.

k

This is an array of twenty one spring stiffness values.

Value(0) = U1U1 [F/L]

Value(1) = U1U2 [F/L]

Value(2) = U2U2 [F/L]

Value(3) = U1U3 [F/L]

Value(4) = U2U3 [F/L]

Value(5) = U3U3 [F/L]

Value(6) = U1R1 [F/rad]

Value(7) = U2R1 [F/rad]

Value(8) = U3R1 [F/rad]

Value(9) = R1R1 [FL/rad]

Value(10) = U1R2 [F/rad]

Value(11) = U2R2 [F/rad]

Value(12) = U3R2 [F/rad]

Value(13) = R1R2 [FL/rad]

Value(14) = R2R2 [FL/rad]

Value(15) = U1R3 [F/rad]

Value(16) = U2R3 [F/rad]

Value(17) = U3R3 [F/rad]

Value(18) = R1R3 [FL/rad]

Value(19) = R2R3 [FL/rad]

Value(20) = R3R3 [FL/rad]

## Remarks

This function retrieves coupled spring stiffness assignments for a point object.

The spring stiffnesses reported are the sum of all springs assigned to the point object. The spring stiffness values are reported in the point local coordinate system.

The function returns zero if the stiffnesses are successfully retrieved, otherwise it returns a nonzero value. If no springs exist at the point object, the function returns a nonzero value.

## VBA Example



```
Sub GetSpringCoupled()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim k() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'assign coupled spring to a point

      ReDim k(20)

      k(2) = 10

      k(17) = 4

      ret = SapModel.PointObj.SetSpringCoupled("3", k)



   'get coupled spring values

      ReDim k(20)

      ret = SapModel.PointObj.GetSpringCoupled("3", k)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpring](GetSpring_%7BPoint_Object%7D.md)

[SetSpring](SetSpring.md)

[SetSpringCoupled](SetSpringCoupled.md)

[DeleteSpring](DeleteSpring_%7BPoint_Object%7D.md)

[IsSpringCoupled](IsSpringCoupled_%7BPoint_Object%7D.md)
