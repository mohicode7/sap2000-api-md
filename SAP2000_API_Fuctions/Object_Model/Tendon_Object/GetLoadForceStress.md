# GetLoadForceStress

## Syntax

SapObject.SapModel.TendonObj.GetLoadForceStress

## VB6 Procedure

Function GetLoadForceStress(ByVal Name As String, ByRef NumberItems As Long, ByRef TendonName() As String, ByRef LoadPat() As String, ByRef JackFrom() As Long, ByRef LoadType() As Long, ByRef Value() As Double, ByRef CurvatureCoeff() As Double, ByRef WobbleCoeff() As Double, ByRef LossAnchorage() As Double, ByRef LossShortening() As Double, ByRef LossCreep() As Double, ByRef LossShrinkage() As Double, ByRef LossSteelRelax() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing tendon object or group, depending on the value of the ItemType item.

NumberItems

The total number of temperature loads retrieved for the specified tendon objects.

TendonName

This is an array that includes the name of the tendon object associated with each temperature load.

LoadPat

This is an array that includes the name of the load pattern associated with each temperature load.

JackFrom

This is an array that includes 1, 2 or 3, indicating how the tendon is jacked.

1 = Tendon jacked from I-End

2 = Tendon jacked from J-End

3 = Tendon jacked from both ends

LoadType

This is an array that includes either 0 or 1, indicating how the type of load.

0 = Force

1 = Stress

Value

This is an array that includes the load value. [F] when LoadType is 0, and [F/L2] when Loadtype is 1

CurvatureCoeff

This is an array that includes the curvature coefficient used when calculating friction losses.

WobbleCoeff

This is an array that includes the wobble coefficient used when calculating friction losses. [1/L]

LossAnchorage

This is an array that includes the anchorage set slip. [L]

LossShortening

This is an array that includes the tendon stress loss due to elastic shortening. [F/L2]

LossCreep

This is an array that includes the tendon stress loss due to creep. [F/L2]

LossShrinkage

This is an array that includes the tendon stress loss due to shrinkage. [F/L2]

LossSteelRelax

This is an array that includes the tendon stress loss due to tendon steel relaxation. [F/L2]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the tendon object specified by the Name item.

If this item is Group, the assignments are retrieved for all tendon objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected tendon objects, and the Name item is ignored.

## Remarks

This function retrieves the force/stress load assignments to tendon objects.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonForceLoad()

   'dimensionvariables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberItems As Long

      Dim TendonName() As String

      Dim LoadPat() As String

      Dim JackFrom() As Long

      Dim LoadType() As Long

      Dim Value() As Double

      Dim CurvatureCoeff() As Double

      Dim WobbleCoeff() As Double

      Dim LossAnchorage() As Double

      Dim LossShortening() As Double

      Dim LossCreep() As Double

      Dim LossShrinkage() As Double

      Dim LossSteelRelax() As Double



   'createSap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'startSap2000 application

      SapObject.ApplicationStart



   'createSapModel object

      Set SapModel = SapObject.SapModel



   'initializemodel

      ret = SapModel.InitializeNewModel



   'createmodel from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'assign tendon force load

      ret = SapModel.TendonObj.SetLoadForceStress("ALL", "DEAD", 1, 0, 100, 0.15, 8.333E-05, 0.25, 3, 5, 7, 5, , Group)



   'get tendon force load

      ret = SapModel.TendonObj.GetLoadForceStress(Name, NumberItems, TendonName, LoadPat, JackFrom, LoadType, Value, CurvatureCoeff, WobbleCoeff, LossAnchorage, LossShortening, LossCreep, LossShrinkage, LossSteelRelax)



   'closeSap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SetLoadForceStress](SetLoadForceStress.md)

[DeleteLoadForceStress](DeleteLoadForceStress_%7BTendon_Object%7D.md)
