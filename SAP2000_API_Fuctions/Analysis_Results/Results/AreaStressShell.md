# AreaStressShell

## Syntax

SapObject.SapModel.Results.AreaStressShell

## VB6 Procedure

Function AreaStressShell(ByVal name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef PointElm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef S11Top() As Double, ByRef S22Top() As Double, ByRef S12Top() As Double, ByRef SMaxTop() As Double, ByRef SMinTop() As Double, ByRef SAngleTop() As Double, ByRef SVMTop() As Double, ByRef S11Bot() As Double, ByRef S22Bot() As Double, ByRef S12Bot() As Double, ByRef SMaxBot() As Double, ByRef SMinBot() As Double, ByRef SAngleBot() As Double, ByRef SVMBot() As Double, ByRef S13Avg() As Double, ByRef S23Avg() As Double, ByRef SMaxAvg() As Double, ByRef SAngleAvg() As Double) As Long

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

PointElm

This is an array that includes the name of the point element where the results are reported.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

S11Top, S22Top, S12Top, S11Bot, S22Bot, S12Bot

The area element internal S11, S22 and S12 stresses, at the top or bottom of the specified area element, at the specified point element location, reported in the area element local coordinate system. [F/L2]

SMaxTop, SMinTop, SMaxBot, SMinBot

The area element maximum and minimum principal stresses, at the top or bottom of the specified area element, at the specified point element location. [F/L2]

SAngleTop, SAngleBot

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the area local 1 axis to the direction of the maximum principal stress, at the top or bottom of the specified area element. [deg]

SVMTop, SVMBot

The area element internal top or bottom Von Mises stress at the specified point element. [F/L2]

S13Avg, S23Avg

The area element average S13 or S23 out-of-plane shear stress at the specified point element. These items are only reported for area elements with properties that allow plate bending behavior. [F/L2]

SMaxAvg

The area element maximum average out-of-plane shear stress. It is equal to the square root of the sum of the squares of S13Avg and S23Avg. This item is only reported for area elements with properties that allow plate bending behavior. [F/L2]

SAngleAvg

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the area local 1 axis to the direction of SMaxAvg. This item is only reported for area elements with properties that allow plate bending behavior. [deg]

## Remarks

This function reports the area stresses for the specified area elements that are assigned shell section properties (not plane or asolid properties). Stresses are reported at each point element associated with the area element.

The function returns zero if the stresses are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetAreaStressesShell()

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

          Dim S11Top() As Double

          Dim S22Top() As Double

          Dim S12Top() As Double

          Dim SMaxTop() As Double

          Dim SMinTop() As Double

          Dim SAngleTop() As Double

          Dim SVMTop() As Double

          Dim S11Bot() As Double

          Dim S22Bot() As Double

          Dim S12Bot() As Double

          Dim SMaxBot() As Double

          Dim SMinBot() As Double

          Dim SAngleBot() As Double

          Dim SVMBot() As Double

          Dim S13Avg() As Double

          Dim S23Avg() As Double

          Dim SMaxAvg() As Double

          Dim SAngleAvg() As Double



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'create model from template

          ret = SapModel.File.NewWall(6, 48, 6, 48)



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")



       'get area stresses for area object "1"

          ret = SapModel.Results.AreaStressShell("1", Object, NumberResults, Obj, Elm, PointElm, LoadCase, StepType, StepNum, S11Top, S22Top, S12Top, SMaxTop, SMinTop, SAngleTop, SVMTop, S11Bot, S22Bot, S12Bot, SMaxBot, SMinBot, SAngleBot, SVMBot, S13Avg, S23Avg, SMaxAvg, SAngleAvg)



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

[AreaJointForceShell](AreaJointForceShell.md)

[AreaStressShellLayered](AreaStressShellLayered.md)

[AreaJointForceShell](AreaJointForceShell.md)
