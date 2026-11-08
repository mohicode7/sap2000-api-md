# ModeShape

## Syntax

SapObject.SapModel.Results.ModeShape

## VB6 Procedure

Function ModeShape(ByVal Name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef U1() As Double, ByRef U2() As Double, ByRef U3() As Double, ByRef R1() As Double, ByRef R2() As Double, ByRef R3() As Double) As Long

## Parameters

Name

The name of an existing point element or group of objects, depending on the value of the ItemTypeElm item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the result request is for the point element corresponding to the point object specified by the Name item.

If this item is Element, the result request is for the point element specified by the Name item.

If this item is GroupElm, the result request is for all point elements directly or indirectly specified in the group specified by the Name item.

If this item is SelectionElm, the result request is for all point elements directly or indirectly selected and the Name item is ignored.

See [Item Type for Elements](Item_Type_for_Elements.md) for more information.

NumberResults

The total number of results returned by the program.

Obj

This is an array that includes the point object name associated with each result, if any. Some results will have no point object associated with them. For these cases this item will be blank.

Elm

This is an array that includes the point element name associated with each result.

LoadCase

This is an array that includes the name of the modal analysis case associated with each result.

StepType

This is an array that includes the step type, if any, for each result. For mode shape results, this is always be Mode. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number for each result. For mode shape results, this is always the mode number. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

U1, U2, U3

These are one dimensional arrays that include the displacement in the point element local 1, 2 and 3 axes directions, respectively, for each result. [L]

R1, R2, R3

These are one dimensional arrays that include the rotation about the point element local 1, 2 and 3 axes, respectively, for each result. [rad]

## Remarks

This function reports the modal displacements (mode shapes) for the specified point elements.

The function returns zero if the displacements are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetModeShapes()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim Obj() As String

          Dim Elm() As String

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim U1() As Double

          Dim U2() As Double

          Dim U3() As Double

          Dim R1() As Double

          Dim R2() As Double

          Dim R3() As Double



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



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("MODAL")



       'get mode shape

          ret = SapModel.Results.ModeShape("ALL", GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)



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

[ModalPeriod](ModalPeriod.md)
