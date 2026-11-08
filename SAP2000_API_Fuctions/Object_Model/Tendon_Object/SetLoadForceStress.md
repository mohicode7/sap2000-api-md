# SetLoadForceStress

## Syntax

SapObject.SapModel.TendonObj.SetLoadForceStress

## VB6 Procedure

Function SetLoadForceStress(ByVal Name As String, ByVal LoadPat As String, ByVal JackFrom As Long, ByVal LoadType As Long, ByVal Value As Double, ByVal CurvatureCoeff As Double, ByVal WobbleCoeff As Double, ByVal LossAnchorage As Double, ByVal LossShortening As Double, ByVal LossCreep As Double, ByVal LossShrinkage As Double, ByVal LossSteelRelax As Double, Optional ByVal Replace As Boolean = True, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing tendon object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

JackFrom

This is 1, 2 or 3, indicating how the tendon is jacked.

1 = Tendon jacked from I-End

2 = Tendon jacked from J-End

3 = Tendon jacked from both ends

LoadType

This is either 0 or 1, indicating how the type of load.

0 = Force

1 = Stress

Value

This is the load value. [F] whenLoadType is 0, and [F/L2] when Loadtype is 1

CurvatureCoeff

The curvature coefficient used when calculating friction losses.

WobbleCoeff

The wobble coefficient used when calculating friction losses. [1/L]

LossAnchorage

The anchorage set slip. [L]

LossShortening

The tendon stress loss due to elastic shortening. [F/L2]

LossCreep

The tendon stress loss due to creep. [F/L2]

LossShrinkage

The tendon stress loss due to shrinkage. [F/L2]

LossSteelRelax

The tendon stress loss due to tendon steel relaxation. [F/L2]

Replace

If this item is True, all previous force/stress loads, if any, assigned to the specified tendon object(s), in the specified load case, are deleted before making the new assignment.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the tendon object specified by the Name item.

If this item is Group, the assignment is made to all tendon objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected tendon objects, and the Name item is ignored.

## Remarks

This function assigns force/stress loads to tendon objects.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignTendonForceLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



   'createSap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'assign tendon force load

      ret = SapModel.TendonObj.SetLoadForceStress("ALL", "DEAD", 1, 0, 100, 0.15, 8.333E-05, 0.25, 3, 5, 7, 5, , Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetLoadForceStress](GetLoadForceStress.md)

[DeleteLoadForceStress](DeleteLoadForceStress_%7BTendon_Object%7D.md)
