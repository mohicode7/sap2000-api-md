# GetSummaryResultsColumn

## Syntax

SapObject.SapModel.DesignConcrete.GetSummaryResultsColumn

## VB6 Procedure

Function GetSummaryResultsColumn(ByVal Name As String, ByRef NumberItems As Long, ByRef FrameName() As String, ByRef MyOption() As Long, ByRef Location() As Double, ByRef PMMCombo() As String, ByRef PMMArea() As Double, ByRef PMMRatio() As Double, ByRef VmajorCombo() As String, ByRef AVmajor() As Double, ByRef VminorCombo() As String, ByRef AVminor() As Double, ByRef ErrorSummary() As String, ByRef WarningSummary() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

NumberItems

The number of frame objects for which results are obtained.

FrameName

This is an array that includes each frame object name for which results are obtained.

MyOption

This is an array that includes 1 or 2, indicating the design option for each frame object.

1 = Check

2 = Design

Location

This is an array that includes the distance from the I-end of the frame object to the location where the results are reported. [L]

PMMCombo

This is an array that includes the name of the design combination for which the controlling PMM ratio or rebar area occurs. A combination name followed by (Sp) indicates that the design loads were obtained by applying special, code-specific multipliers to all or part of the specified design load combination, or that the design was based on the capacity of other objects (or other design locations for the same object).

PMMArea

This is an array that includes the total longitudinal rebar area required for the axial force plus biaxial moment (PMM) design at the specified location. [L2]

This item applies only when MyOption = 2 (design).

PMMRatio

This is an array that includes the axial force plus biaxial moment (PMM) stress ratio at the specified location.

This item applies only when MyOption = 1 (check).

VmajorCombo

This is an array that includes the name of the design combination for which the controlling major shear occurs.

AVmajor

This is an array that includes the required area of transverse shear reinforcing per unit length along the frame object for major shear at the specified location. [L2/L]

VminorCombo

This is an array that includes the name of the design combination for which the controlling minor shear occurs.

AVminor

This is an array that includes the required area of transverse shear reinforcing per unit length along the frame object for minor shear at the specified location. [L2/L]

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

## VBA Example



```
Sub GetConcreteColumnDesignSummaryResults()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberItems As Long

      Dim FrameName() As String

      Dim MyOption() As Long

      Dim Location() As Double

      Dim PMMCombo() As String

      Dim PMMArea() As Double

      Dim PMMRatio() As Double

      Dim VmajorCombo() As String

      Dim AVmajor() As Double

      Dim VminorCombo() As String

      Dim AVminor() As Double

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



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start concrete design

      ret = SapModel.DesignConcrete.StartDesign



   'get summary result data

      ret = SapModel.DesignConcrete.GetSummaryResultsColumn("1", NumberItems, FrameName, MyOption, Location, PMMCombo, PMMArea, PMMRatio, VmajorCombo, AVmajor, VminorCombo, AVminor, ErrorSummary, WarningSummary)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetSummaryResultsBeam](GetSummaryResultsBeam.md)

[GetSummaryResultsJoint](GetSummaryResultsJoint.md)
