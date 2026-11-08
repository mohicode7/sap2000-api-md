# GetSpring

## Syntax

SapObject.SapModel.PointObj.GetSpring

## VB6 Procedure

Function GetSpring(ByVal Name As String, ByRef k() As Double) As Long

## Parameters

Name

The name of an existing point object.

k

This is an array of six spring stiffness values.

Value(0) = U1 [F/L]

Value(1) = U2 [F/L]

Value(2) = U3 [F/L]

Value(3) = R1 [FL/rad]

Value(4) = R2 [FL/rad]

Value(5) = R3 [FL/rad]

## Remarks

This function retrieves uncoupled spring stiffness assignments for a point object, that is, it retrieves the diagonal terms in the 6x6 spring matrix for the point object.

The spring stiffnesses reported are the sum of all springs assigned to the point object. The spring stiffness values are reported in the point local coordinate system.

The function returns zero if the stiffnesses are successfully retrieved, otherwise it returns a nonzero value. If no springs exist at the point object, the function returns a nonzero value.

## VBA Example



```
Sub GetSpring()

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



   'assign spring to a point

      ReDim k(5)

      k(2) = 10

      ret = SapModel.PointObj.SetSpring("3", k)



   'get spring values

      ReDim k(5)

      ret = SapModel.PointObj.GetSpring("3", k)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpringCoupled](GetSpringCoupled_%7BPoint_Object%7D.md)

[SetSpring](SetSpring.md)

[SetSpringCoupled](SetSpringCoupled.md)

[DeleteSpring](DeleteSpring_%7BPoint_Object%7D.md)

[IsSpringCoupled](IsSpringCoupled_%7BPoint_Object%7D.md)
