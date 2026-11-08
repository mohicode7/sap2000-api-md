# SetGroup

## Syntax

SapObject.SapModel.GroupDef.SetGroup

## VB6 Procedure

Function SetGroup(ByVal Name As String, Optional ByVal color As Long = -1, Optional ByVal SpecifiedForSelection As Boolean = True, Optional ByVal SpecifiedForSectionCutDefinition As Boolean = True, Optional ByVal SpecifiedForSteelDesign As Boolean = True, Optional ByVal SpecifiedForConcreteDesign As Boolean = True, Optional ByVal SpecifiedForAluminumDesign As Boolean = True, Optional ByVal SpecifiedForColdFormedDesign As Boolean = True, Optional ByVal SpecifiedForStaticNLActiveStage As Boolean = True, Optional ByVal SpecifiedForBridgeResponseOutput As Boolean = True, Optional ByVal SpecifiedForAutoSeismicOutput As Boolean = False, Optional ByVal SpecifiedForAutoWindOutput As Boolean = False, Optional ByVal SpecifiedForMassAndWeight As Boolean = True) As Long

## Parameters

Name

This is the name of a group.  If this is the name of an existing group,  that group is modified, otherwise a new group is added.

color

The display color for the group specified as a Long. If this value is input as –1, the program automatically selects a display color for the group.

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

The function returns zero if the group data is successfully set, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetGroupData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'define new group

      ret = SapModel.GroupDef.SetGroup("Group1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetGroup](GetGroup_%7BGroup%7D.md)
