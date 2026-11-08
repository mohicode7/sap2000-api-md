# PanelZoneForce

## Syntax

SapObject.SapModel.Results.PanelZoneForce

## VB6 Procedure

Function PanelZoneForce(ByVal Name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Elm() As String, ByRef PointElm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef P() As Double, ByRef V2() As Double, ByRef V3() As Double, ByRef T() As Double, ByRef M2() As Double, ByRef M3() As Double) As Long

## Parameters

Name

The name of an existing point object, point element, or group of objects, depending on the value of the ItemTypeElm item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the result request is for the panel zone (link) element corresponding to the panel zone assignment to the point object specified by the Name item.

If this item is Element, the result request is for the panel zone (link) element specified by the Name item.

If this item is GroupElm, the result request is for all panel zone (link) elements directly or indirectly specified in the group specified by the Name item.

If this item is SelectionElm, the result request is for panel zone (link) elements directly or indirectly selected and the Name item is ignored.

For GroupElm and SelectionElm a panel zone (link) element may be indirectly specified through point objects that have panel zone assignments.

NumberResults

The total number of results returned by the program.

Elm

This is an array that includes the panel zone (link) element name associated with each result.

PointElm

This is an array that includes the point element name associated with each result.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

P

This is an array that includes the panel zone (link) axial force (in the link local 1-axis direction) at the specified point element. [F]

V2, V3

These are one dimensional arrays that include the panel zone (link) shear force components in the link element local axes directions. [F]

T

This is an array that includes the panel zone (link) torsion (about the link local 1-axis) at the specified point element. [FL]

M2, M3

These are one dimensional arrays that include the panel zone (link) moment components about the link element local axes. [FL]

## Remarks

This function reports the panel zone (link) forces at the point elements at the ends of the specified panel zone (link) elements.

The function returns zero if the forces are successfully recovered, otherwise it returns a nonzero value.

See Analysis Result Remarks for more information.

## VBA Example



```
Sub GetPanelZoneForce()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim Elm() As String

          Dim PointElm() As String

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim P() As Double

          Dim V2() As Double

          Dim V3() As Double

          Dim T() As Double

          Dim M2() As Double

          Dim M3() As Double



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



       'assign panel zone to point object "3"

          ret = SapModel.PointObj.SetPanelZone("3", 1, 2, 0, 0, "", 0, 0, 0)



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")



       'get panel zone force for point object "3"

          ret = SapModel.Results.PanelZoneForce("3", ObjectElm, NumberResults, Elm, PointElm, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)



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

[PanelZoneDeformation](PanelZoneDeformation.md)
