# SetBeam

## Syntax

SapObject.SapModel.ConstraintDef.SetBeam

## VB6 Procedure

Function SetBeam(ByVal Name As String, Optional ByVal Axis As eConstraintAxis = AutoAxis, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of a constraint.

Axis

This is one of the following items from the eConstraintAxis enumeration. It specifies the axis in the specified coordinate system that is parallel to the axis of the constraint. If AutoAxis is specified, the axis of the constraint is automatically determined from the joints assigned to the constraint.

X = 1

Y = 2

Z = 3

AutoAxis = 4

CSys

The name of the coordinate system in which the constraint is defined.

## Remarks

This function defines a Beam constraint. If the specified name is not used for a constraint, a new constraint is defined using the specified name. If the specified name is already used for another Beam constraint, the definition of that constraint is modified. If the specified name is already used for some constraint that is not a Beam constraint, an error is returned.

The function returns zero if the constraint data is successfully added or modified, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetBeamConstraint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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

      ret = SapModel.ConstraintDef.SetBeam("Beam1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetBeam](GetBeam.md)
