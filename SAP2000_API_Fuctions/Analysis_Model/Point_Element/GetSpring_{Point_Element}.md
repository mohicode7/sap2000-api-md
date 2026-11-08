# GetSpring

## Syntax

SapObject.SapModel.PointElm.GetSpring

## VB6 Procedure

Function GetSpring(ByVal Name As String, ByRef k() As Double) As Long

## Parameters

Name

The name of an existing point element.

k

This is an array of six spring stiffness values.

Value(0) = U1 [F/L]

Value(1) = U2 [F/L]

Value(2) = U3 [F/L]

Value(3) = R1 [FL/rad]

Value(4) = R2 [FL/rad]

Value(5) = R3 [FL/rad]

## Remarks

This function retrieves uncoupled spring stiffness assignments for a point element; that is, it retrieves the diagonal terms in the 6x6 spring matrix for the point element.

The spring stiffnesses reported are the sum of all springs assigned to the point element either directly or indirectly through line, area and solid spring assignments. The spring stiffness values are reported in the point local coordinate system.

The function returns zero if the stiffnesses are successfully retrieved, otherwise it returns a nonzero value. If no springs exist at the point element, the function returns a nonzero value.

## VBA Example



```
Sub GetPointElmSpring()

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



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get spring values at point element

      ReDim k(5)

      ret = SapModel.PointElm.GetSpring("3", k)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpringCoupled](GetSpringCoupled_%7BPoint_Element%7D.md)

[IsSpringCoupled](IsSpringCoupled_%7BPoint_Element%7D.md)
