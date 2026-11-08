# AreaStrainShell

## Syntax

SapObject.SapModel.Results.AreaStrainShell

## VB6 Procedure

Function AreaStrainShell(ByVal name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef PointElm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef E11Top() As Double, ByRef E22Top() As Double, ByRef G12Top() As Double, ByRef EMaxTop() As Double, ByRef EMinTop() As Double, ByRef EAngleTop() As Double, ByRef EVMTop() As Double, ByRef E11Bot() As Double, ByRef E22Bot() As Double, ByRef G12Bot() As Double, ByRef EMaxBot() As Double, ByRef EMinBot() As Double, ByRef EAngleBot() As Double, ByRef EVMBot() As Double, ByRef G13Avg() As Double, ByRef G23Avg() As Double, ByRef GMaxAvg() As Double, ByRef GAngleAvg() As Double) As Long

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

E11Top, E22Top, G12Top, E11Bot, E22Bot, G12Bot

The area element internal E11, E22 and G12 strains, at the top or bottom of the specified area element, at the specified point element location, reported in the area element local coordinate system.

EMaxTop, EMinTop, EMaxBot, EMinBot

The area element maximum and minimum principal strains, at the top or bottom of the specified area element, at the specified point element location.

EAngleTop, EAngleBot

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the area local 1 axis to the direction of the maximum principal strain, at the top or bottom of the specified area element. [deg]

EVMTop, EVMBot

The area element internal top or bottom Von Mises strain at the specified point element.

G13Avg, G23Avg

The area element average G13 or G23 out-of-plane shear strain at the specified point element. These items are only reported for area elements with properties that allow plate bending behavior.

GMaxAvg

The area element maximum average out-of-plane shear strain. It is equal to the square root of the sum of the squares of G13Avg and G23Avg. This item is only reported for area elements with properties that allow plate bending behavior.

GAngleAvg

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the area local 1 axis to the direction of GMaxAvg. This item is only reported for area elements with properties that allow plate bending behavior. [deg]

## Remarks

This function reports the area strains for the specified area elements that are assigned shell section properties (not plane or asolid properties). Strains are reported at each point element associated with the area element.

The function returns zero if the strains are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetAreaStrainsShell()

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

          Dim E11Top() As Double

          Dim E22Top() As Double

          Dim G12Top() As Double

          Dim EMaxTop() As Double

          Dim EMinTop() As Double

          Dim EAngleTop() As Double

          Dim EVMTop() As Double

          Dim E11Bot() As Double

          Dim E22Bot() As Double

          Dim G12Bot() As Double

          Dim EMaxBot() As Double

          Dim EMinBot() As Double

          Dim EAngleBot() As Double

          Dim EVMBot() As Double

          Dim G13Avg() As Double

          Dim G23Avg() As Double

          Dim GMaxAvg() As Double

          Dim GAngleAvg() As Double



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



       'get area strains for area object "1"

          ret = SapModel.Results.AreaStrainShell("1", Object, NumberResults, Obj, Elm, PointElm, LoadCase, StepType, StepNum, E11Top, E22Top, G12Top, EMaxTop, EMinTop, EAngleTop, EVMTop, E11Bot, E22Bot, G12Bot, EMaxBot, EMinBot, EAngleBot, EVMBot, G13Avg, G23Avg, GMaxAvg, GAngleAvg)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 20.1.0.

## See Also

[AreaStrainShellLayered](AreaStrainShellLayered.md)
