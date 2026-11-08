# GetSummaryResults

## Syntax

SapObject.SapModel.DesignColdFormed.GetSummaryResults

## VB6 Procedure

Function GetSummaryResults(ByVal Name As String, ByRef NumberItems As Long, ByRef FrameName() As String, ByRef Ratio() As Double, ByRef RatioType() As Long, ByRef Location() As Double, ByRef ComboName() As String, ByRef ErrorSummary() As String, ByRef WarningSummary() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

NumberItems

The number of frame objects for which results are obtained.

FrameName

This is an array that includes each frame object name for which results are obtained.

Ratio

This is an array that includes the controlling stress or capacity ratio for each frame object.

RatioType

This is an array that includes 1, 3 or 4, indicating the controlling stress or capacity ratio type for each frame object.

1 = PMM

3 = Major shear

4 = Minor shear

Location

This is an array that includes the distance from the I-end of the frame object to the location where the controlling stress or capacity ratio occurs. [L]

ComboName

This is an array that includes the name of the design combination for which the controlling stress or capacity ratio occurs.

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

This function retrieves summary results for cold formed design.

The function returns zero if the results are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetColdFormedDesignSummaryResults()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberItems As Long

      Dim FrameName() As String

      Dim Ratio() As Double

      Dim RatioType() As Long

      Dim Location() As Double

      Dim ComboName() As String

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



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section property

      ret = SapModel.PropFrame.SetColdC("CdC", Name , 9, 3, 0.06, 0.25, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 72, 2, 144, True, "CdC", "CdC")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start cold formed design

      ret = SapModel.DesignColdFormed.StartDesign



   'get summary result data

      ret = SapModel.DesignColdFormed.GetSummaryResults("8", NumberItems, FrameName, Ratio, RatioType, Location, ComboName, ErrorSummary, WarningSummary)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also
