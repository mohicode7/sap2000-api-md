# GetSectionCutSelectedForOutput

## Syntax

SapObject.SapModel.Results.Setup.GetSectionCutSelectedForOutput

## VB6 Procedure

Function GetSectionCutSelectedForOutput (ByVal Name As String, ByRef Selected As Boolean) AsLong

## Parameters

Name

The name of a defined section cut.

Selected

This item is True if the section cut is to be selected for output, or False if the section cut is not selected for output.

## Remarks

This function retrieves whether a defined section cut is selected for output.

The function returns 0 if the selected flag is successfully retrieved, otherwise it returns nonzero.

Please note that all section cuts are, by default, selected for output when they are created.

## VBA Example



```
Sub SelectAllSectionCuts()
'dimension variables
Dim SapObject As Sap2000v16.SapObject
Dim SapModel As cSapModel
Dim ret As Long
Dim Selected As Boolean
 
'create Sap2000 object
Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")
 
'start Sap2000 application
SapObject.ApplicationStart
 
'create SapModel object
Set SapModel = SapObject.SapModel
 
'initialize model
ret = SapModel.InitializeNewModel
 
'create model from template
ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 288)
 
'define new group
ret = SapModel.GroupDef.SetGroup("Group1")
 
'add objects to group
ret = SapModel.PointObj.SetGroupAssign("1", "Group1")
ret = SapModel.FrameObj.SetGroupAssign("1", "Group1")
 
'define section cut
ret = SapModel.SectCut.SetByGroup("SCut1", "Group1", 1)
 
'Check if the section cut is selected for output
ret = SapModel.Results.Setup.GetSectionCutSelectedForOutput("SCut1", Selected)
 
'close Sap2000
SapObject.ApplicationExit False
Set SapModel = Nothing
Set SapObject = Nothing
 
End Sub
 
```



## Release Notes

Initial release in version 16.0.2.

## See Also

[SelectAllSectionCutsForOutput](SelectAllSectionCutsForOutput.md)

[SetSectionCutSelectedForOutput](SetSectionCutSelectedForOutput.md)
