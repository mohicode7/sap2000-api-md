# AreaStrainShellLayered

## Syntax

SapObject.Sap2000.Results.AreaStrainShellLayered

## VB6 Procedure

Function AreaStrainShellLayered(ByVal name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef Layer() As String, ByRef IntPtNum() As Long, ByRef IntPtLoc() As Double, ByRef PointElm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef E11() As Double, ByRef E22() As Double, ByRef G12() As Double, ByRef EMax() As Double, ByRef EMin() As Double, ByRef EAngle() As Double, ByRef EVM() As Double, ByRef G13Avg() As Double, ByRef G23Avg() As Double, ByRef GMaxAvg() As Double, ByRef GAngleAvg() As Double) As Long

## Parameters

Name

The name of an existing area object, area element or group of objects, depending on the value of the ItemTypeElm item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the result request is for the area elements corresponding to the area object specified by the Name item.

If this item is Element, the result request is for the area element specified by the Name item.

If this item is GroupElm, the result request is for the area elements corresponding to all area objects included in the group specified by the Name item.

If this item is SelectionElm, the result request is for area elements corresponding to all selected area objects and the Name item is ignored.

NumberResults

The total number of results returned by the program.

Obj

This is an array that includes the area object name associated with each result, if any.

Elm

This is an array that includes the area element name associated with each result.

Layer

This is an array that includes the layer name associated with each result.

IntPtNum

This is an array that includes the integration point number within the specified layer of the area element.

IntPtLoc

This is an array that includes the integration point relative location within the specified layer of the area element. The location is between -1 (bottom of layer) and +1 (top of layer), inclusive. The midheight of the layer is at a value of 0.

PointElm

This is an array that includes the name of the point element where the results are reported.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

E11, E22, G12

The area element internal E11, E22 and G12 strains, at the specified point element location, for the specified layer and layer integration point, reported in the area element local coordinate system.

EMax, EMin

The area element maximum and minimum principal strains, at the specified point element location, for the specified layer and layer integration point.

EAngle

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the area local 1 axis to the direction of the maximum principal strain. [deg]

EVM

The area element internal Von Mises strain at the specified point element location, for the specified layer and layer integration point.

G13Avg, G23Avg

The area element average G13 or G23 out-of-plane shear strain at the specified point element location, for the specified layer and layer integration point.

GMaxAvg

The area element maximum average out-of-plane shear strain for the specified layer and layer integration point. It is equal to the square root of the sum of the squares of G13Avg and G23Avg.

GAngleAvg

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the area local 1 axis to the direction of GMaxAvg. [deg]

## Remarks

This function reports the area strains for the specified area elements that are assigned layered shell section properties.

The function returns zero if the strains are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Similar to AreaStrainShell



       'get area strains for area object "1"

          ret = SapModel.Results.AreaStrainShellLayered("1", Object, NumberResults, Obj, Elm, Layer, IntPtNum, IntPtLoc, PointElm, LoadCase, StepType, StepNum, E11, E22, G12, EMax, EMin, EAngle, EVM, G13Avg, G23Avg, GMaxAvg, GAngleAvg )
  
```



## Release Notes

Initial release in version 20.1.0.

## See Also

[AreaStrainShell](AreaStrainShell.md)
