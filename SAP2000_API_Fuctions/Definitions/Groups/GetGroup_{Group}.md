# GetGroup

## Syntax

SapObject.SapModel.GroupDef.GetGroup

## VB6 Procedure

Function GetGroup(ByVal Name As String, ByRef color As Long, ByRef SpecifiedForSelection As Boolean,ByRef SpecifiedForSectionCutDefinition As Boolean, ByRef SpecifiedForSteelDesign As Boolean, ByRef SpecifiedForConcreteDesign As Boolean, ByRef SpecifiedForAluminumDesign As Boolean, ByRef SpecifiedForColdFormedDesign As Boolean, ByRef SpecifiedForStaticNLActiveStage As Boolean, ByRef SpecifiedForBridgeResponseOutput As Boolean, ByRef SpecifiedForAutoSeismicOutput As Boolean, ByRef SpecifiedForAutoWindOutput As Boolean, ByRef SpecifiedForMassAndWeight As Boolean) As Long

## Parameters

Name

The name of an existing group.

color

The display color for the group specified as a Long.

SpecifiedForSelection

This item is True if the group is specified to be used for selection; otherwise it is False.

SpecifiedForSectionCutDefinition

This item is True if the group is specified to be used for defining section cuts; otherwise it is False.

SpecifiedForSteelDesign

This item is True if the group is specified to be used for defining steel frame design groups; otherwise it is False.

SpecifiedForConcreteDesign

This item is True if the group is specified to be used for defining concrete frame design groups; otherwise it is False.

SpecifiedForAluminumDesign

This item is True if the group is specified to be used for defining aluminum frame design groups; otherwise it is False.

SpecifiedForColdFormedDesign

This item is True if the group is specified to be used for defining cold formed frame design groups; otherwise it is False.

SpecifiedForStaticNLActiveStage

This item is True if the group is specified to be used for defining stages for nonlinear static analysis; otherwise it is False.

SpecifiedForBridgeResponseOutput

This item is True if the group is specified to be used for reporting bridge response output; otherwise it is False.

SpecifiedForAutoSeismicOutput

This item is True if the group is specified to be used for reporting auto seismic loads; otherwise it is False.

SpecifiedForAutoWindOutput

This item is True if the group is specified to be used for reporting auto wind loads; otherwise it is False.

SpecifiedForMassAndWeight

This item is True if the group is specified to be used for reporting group masses and weight; otherwise it is False.

## Remarks

The function returns zero if the group data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetGroupData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim Color as long

      Dim SpecifiedForSelection as Boolean

      Dim SpecifiedForSectionCutDefinition as Boolean

      Dim SpecifiedForSteelDesign as Boolean

      Dim SpecifiedForConcreteDesign as Boolean

      Dim SpecifiedForAluminumDesign as Boolean

      Dim SpecifiedForColdFormedDesign as Boolean

      Dim SpecifiedForStaticNLActiveStage as Boolean

      Dim SpecifiedForBridgeResponseOutput as Boolean

      Dim SpecifiedForAutoSeismicOutput as Boolean

      Dim SpecifiedForAutoWindOutput as Boolean

      Dim SpecifiedForMassAndWeight as Boolean



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



   'get group data

      ret = SapModel.GroupDef.GetGroup("ALL", Color, SpecifiedForSelection, SpecifiedForSectionCutDefinition, SpecifiedForSteelDesign, SpecifiedForConcreteDesign, SpecifiedForAluminumDesign, SpecifiedForColdFormedDesign, SpecifiedForStaticNLActiveStage, SpecifiedForBridgeResponseOutput, SpecifiedForAutoSeismicOutput, SpecifiedForAutoWindOutput, SpecifiedForMassAndWeight)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetGroup](SetGroup_%7BGroup%7D.md)
