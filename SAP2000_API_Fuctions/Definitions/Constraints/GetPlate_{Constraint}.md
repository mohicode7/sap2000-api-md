# GetPlate

## Syntax

SapObject.SapModel.ConstraintDef.GetPlate

## VB6 Procedure

Function GetPlate(ByVal Name As String, ByRef Axis As eConstraintAxis, ByRef CSys As String) As Long

## Parameters

Name

The name of an existing constraint.

Axis

This is one of the following items from the eConstraintAxis enumeration. It specifies the axis in the specified coordinate system that is perpendicular to the plane of the constraint. If AutoAxis is specified, the axis of the constraint is automatically determined from the joints assigned to the constraint.

X = 1

Y = 2

Z = 3

AutoAxis = 4

CSys

The name of the coordinate system in which the constraint is defined.

## Remarks

The function returns the definition for the specified constraint.

The function returns zero if the constraint data is successfully obtained, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPlateConstraint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ConstraintType As eConstraintType

      Dim Axis As eConstraintAxis

      Dim CSys as String



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



   'define new constraint

      ret = SapModel.ConstraintDef.SetPlate("Plate1")



   'get constraint data

      ret = SapModel.ConstraintDef.GetConstraintType("Plate1", ConstraintType)

      If ConstraintType = CONSTRAINT_PLATE Then

         ret = SapModel.ConstraintDef.GetPlate("Plate1", Axis, CSys)

      End If



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetPlate](SetPlate_%7BConstraint%7D.md)
