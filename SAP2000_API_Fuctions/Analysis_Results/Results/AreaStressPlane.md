# AreaStressPlane

## Syntax

SapObject.Sap2000.Results.AreaStressPlane

## VB6 Procedure

Function AreaStressPlane(ByVal name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef PointElm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef S11() As Double, ByRef S22() As Double, ByRef S33() As Double, ByRef S12() As Double, ByRef SMax() As Double, ByRef SMin() As Double, ByRef SAngle() As Double, ByRef SVM() As Double) As Long

## Parameters

Name

The name of an existing area object, area element or group of objects, depending on the value of the ItemTypeElm item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the result request is for the plane elements corresponding to the area object specified by the Name item.

If this item is Element, the result request is for the plane element specified by the Name item.

If this item is GroupElm, the result request is for the plane elements corresponding to all area objects included in the group specified by the Name item.

If this item is SelectionElm, the result request is for plane elements corresponding to all selected area objects, and the Name item is ignored.

NumberResults

The total number of results returned by the program.

Obj

This is an array that includes the area object name associated with each result, if any.

Elm

This is an array that includes the plane element name associated with each result.

PointElm

This is an array that includes the name of the point element where the results are reported.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

S11, S22, S33, S12

The plane element internal S11, S22, S33 and S12 stresses, at the specified point element location, reported in the area element local coordinate system. [F/L2]

SMax, SMin

The plane element maximum and minimum principal stresses at the specified point element location. [F/L2]

SAngle

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the plane element local 1 axis to the direction of the maximum principal stress. [deg]

SVM

The plane element internal Von Mises stress at the specified point element. [F/L2]

## Remarks

This function reports the stresses for the specified plane elements that are assigned plane or asolid section properties (not shell properties).

The function returns zero if the stresses are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetPlaneStresses()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim Obj() As String

          Dim Elm() As String

          Dim PointElm() As String

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim S11() As Double

          Dim S22() As Double

          Dim S33() As Double

          Dim S12() As Double

          Dim SMax() As Double

          Dim SMin() As Double

          Dim SAngle() As Double

          Dim SVM() As Double



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'open existing model

          ret = SapModel.File.OpenFile("C:\SapAPI\Example 3-001-incomp.sdb")



       'run analysis

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("MEMBRANE")



       'get plane stresses for area object "1"

          ret = SapModel.Results.AreaStressPlane("1", ObjectElm, NumberResults, Obj, Elm, PointElm, LoadCase, StepType, StepNum, S11, S22, S33, S12, SMax, SMin, SAngle, SVM)



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

[AreaJointForcePlane](AreaJointForcePlane.md)
