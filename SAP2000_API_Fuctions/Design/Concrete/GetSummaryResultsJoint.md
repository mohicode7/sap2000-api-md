# GetSummaryResultsJoint

## Syntax

SapObject.SapModel.DesignConcrete.GetSummaryResultsJoint

## VB6 Procedure

Function GetSummaryResultsJoint(ByVal Name As String, ByRef NumberItems As Long, ByRef FrameName() As String, ByRef LCJSRatioMajor() As String, ByRef JSRatioMajor() As Double, ByRef LCJSRatioMinor() As String, ByRef JSRatioMinor() As Double, ByRef LCBCCRatioMajor() As String, ByRef BCCRatioMajor() As Double, ByRef LCBCCRatioMinor() As String, ByRef BCCRatioMinor() As Double, ByRef ErrorSummary() As String, ByRef WarningSummary() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

NumberItems

The number of frame objects for which results are obtained.

FrameName

This is an array that includes each frame object name for which results are obtained.

LCJSRatioMajor

This is an array that includes the name of the design combination for which the controlling joint shear ratio associated with the column major axis occurs. A combination name followed by (Sp) indicates that the design loads were obtained by applying special, code-specific multipliers to all or part of the specified design load combination, or that the design was based on the capacity of other objects (or other design locations for the same object).

JSRatioMajor

This is an array that includes the joint shear ratio associated with the column major axis. This is the joint shear divided by the joint shear capacity.

LCJSRatioMinor

This is an array that includes the name of the design combination for which the controlling joint shear ratio associated with the column minor axis occurs. A combination name followed by (Sp) indicates that the design loads were obtained by applying special, code-specific multipliers to all or part of the specified design load combination, or that the design was based on the capacity of other objects (or other design locations for the same object).

JSRatioMinor

This is an array that includes the joint shear ratio associated with the column minor axis. This is the joint shear divided by the joint shear capacity.

LCBCCRatioMajor

This is an array that includes the name of the design combination for which the controlling beam/column capacity ratio associated with the column major axis occurs. A combination name followed by (Sp) indicates that the design loads were obtained by applying special, code-specific multipliers to all or part of the specified design load combination, or that the design was based on the capacity of other objects (or other design locations for the same object).

BCCRatioMajor

This is an array that includes the beam/column capacity ratio associated with the column major axis. This is the sum of the column capacities divided by the sum of the beam capacities at the top of the specified column.

LCBCCRatioMinor

This is an array that includes the name of the design combination for which the controlling beam/column capacity ratio associated with the column minor axis occurs. A combination name followed by (Sp) indicates that the design loads were obtained by applying special, code-specific multipliers to all or part of the specified design load combination, or that the design was based on the capacity of other objects (or other design locations for the same object).

BCCRatioMinor

This is an array that includes the beam/column capacity ratio associated with the column minor axis. This is the sum of the column capacities divided by the sum of the beam capacities at the top of the specified column.

ErrorSummary

This is an array that includes the design error messages for the frame object, if any.

WarningSummary

This is an array that includes the design warning messages for the frame object, if any.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the design results are retrieved for the frame object specified by the Name item.

If this item is Group, the design results are retrieved for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, the design results are retrieved for all selected frame objects, and the Name item is ignored.

## Remarks

This function retrieves summary results for concrete design.

The function returns zero if the results are successfully retrieved; otherwise it returns a nonzero value.

Note that joint design is only included for some codes.

## VBA Example



```
Sub GetConcreteJointDesignSummaryResults()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberItems As Long

      Dim FrameName() As String

      Dim LCJSRatioMajor() As String

      Dim JSRatioMajor() As Double

      Dim LCJSRatioMinor() As String

      Dim JSRatioMinor() As Double

      Dim LCBCCRatioMajor() As String

      Dim BCCRatioMajor() As Double

      Dim LCBCCRatioMinor() As String

      Dim BCCRatioMinor() As Double

      Dim ErrorSummary() As String

      Dim WarningSummary() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'add ASTM A706 rebar material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A706)



   'create new concrete frame section properties

      ret = SapModel.PropFrame.SetRectangle("COL", "4000Psi", 20, 20)

      ret = SapModel.PropFrame.SetRectangle("BEAM", "4000Psi", 20, 12)

      ret = SapModel.PropFrame.SetRebarBeam("BEAM", Name, Name, 2, 2, 2, 2, 2, 2)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "BEAM", "COL")



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("EQX", LTYPE_QUAKE)



   'assign IBC2003 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetIBC2003("EQX", 1, 0.05, 1, 0.035, 0, False, 0, 0, 1, 1, 3, 1, 0.4, 0, 0, 8, 3, 5.5)



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("ACI 318-02")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start concrete design

      ret = SapModel.DesignConcrete.StartDesign



   'get summary result data

      ret = SapModel.DesignConcrete.GetSummaryResultsJoint("3", NumberItems, FrameName, LCJSRatioMajor, JSRatioMajor, LCJSRatioMinor, JSRatioMinor, LCBCCRatioMajor, BCCRatioMajor, LCBCCRatioMinor, BCCRatioMinor, ErrorSummary, WarningSummary)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetSummaryResultsBeam](GetSummaryResultsBeam.md)

[GetSummaryResultsColumn](GetSummaryResultsColumn.md)
