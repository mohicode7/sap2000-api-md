# SolidStress

## Syntax

SapObject.SapModel.Results.SolidStress

## VB6 Procedure

Function SolidStress(ByVal name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef PointElm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef S11() As Double, ByRef S22() As Double, ByRef S33() As Double, ByRef S12() As Double, ByRef S13() As Double, ByRef S23() As Double, ByRef SMax() As Double, ByRef SMid() As Double, ByRef SMin() As Double, ByRef SVM() As Double, ByRef DirCosMax1() As Double, ByRef DirCosMax2() As Double, ByRef DirCosMax3() As Double, ByRef DirCosMid1() As Double, ByRef DirCosMid2() As Double, ByRef DirCosMid3() As Double, ByRef DirCosMin1() As Double, ByRef DirCosMin2() As Double, ByRef DirCosMin3() As Double) As Long

## Parameters

Name

The name of an existing solid object, solid element, or group of objects, depending on the value of the ItemTypeElm item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the result request is for the solid elements corresponding to the solid object specified by the Name item.

If this item is Element, the result request is for the solid element specified by the Name item.

If this item is GroupElm, the result request is for the solid elements corresponding to all solid objects included in the group specified by the Name item.

If this item is SelectionElm, the result request is for solid elements corresponding to all selected solid objects and the Name item is ignored.

NumberResults

The total number of results returned by the program.

Obj

This is an array that includes the solid object name associated with each result, if any.

Elm

This is an array that includes the solid element name associated with each result.

PointElm

This is an array that includes the name of the point element where the results are reported.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

S11, S22, S33, S12, S13, S23

The solid element internal S11, S22, S33, S12, S13 and S23 stresses at the specified point element location, reported in the solid element local coordinate system. [F/L2]

SMax, SMid, SMin

The solid element maximum, middle and minimum principal stresses at the specified point element location. [F/L2]

SVM

The solid element internal Von Mises stress at the specified point element location. [F/L2]

DirCosMax1, DirCosMax2, DirCosMax3

These are three direction cosines defining the orientation of the maximum principal stress with respect to the solid element local axes.

DirCosMid1, DirCosMid2, DirCosMid3

These are three direction cosines defining the orientation of the middle principal stress with respect to the solid element local axes.

DirCosMin1, DirCosMin2, DirCosMin3

These are three direction cosines defining the orientation of the minimum principal stress with respect to the solid element local axes.

## Remarks

This function reports the stresses for the specified solid elements. Stresses are reported at each point element associated with the solid element.

The function returns zero if the stresses are successfully recovered, otherwise it returns a nonzero value.

See Analysis Results Remarks for more information.

## VBA Example



```
Sub GetSolidStresses()

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

          Dim S13() As Double

          Dim S23() As Double

          Dim SMax() As Double

          Dim SMid() As Double

          Dim SMin() As Double

          Dim SVM() As Double

          Dim DirCosMax1() As Double

          Dim DirCosMax2() As Double

          Dim DirCosMax3() As Double

          Dim DirCosMid1() As Double

          Dim DirCosMid2() As Double

          Dim DirCosMid3() As Double

          Dim DirCosMin1() As Double

          Dim DirCosMin2() As Double

          Dim DirCosMin3() As Double



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'create a solid model from template

          ret = SapModel.File.NewSolidBlock(20, 50, 20)



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")



       'get solid stresses for solid object "1"

          ret = SapModel.Results.SolidStress("1", ObjectElm, NumberResults, Obj, Elm, PointElm, LoadCase, StepType, StepNum, S11, S22, S33, S12, S13, S23, SMax, SMid, SMin, SVM, DirCosMax1, DirCosMax2, DirCosMax3, DirCosMid1, DirCosMid2, DirCosMid3, DirCosMin1, DirCosMin2, DirCosMin3)



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

[SolidJointForce](SolidJointForce.md)
