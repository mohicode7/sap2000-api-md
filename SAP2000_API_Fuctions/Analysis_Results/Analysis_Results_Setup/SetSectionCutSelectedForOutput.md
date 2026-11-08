# SetSectionCutSelectedForOutput

## Syntax

SapObject.SapModel.Results.Setup.SetSectionCutSelectedForOutput

## VB6 Procedure

Function SetSectionCutSelectedForOutput (ByVal Name As String, ByVal Selected As Boolean) AsLong

## Parameters

Name

The name of a defined section cut.

Selected

This item is True if the section cut is to be selected for output, or False if no section cut should not be selected for output.

## Remarks

This function selects or deselects a defined section cut for output.

The function returns 0 if the selected flag is successfully set, otherwise it returns nonzero.

Please note that all section cuts are, by default, selected for output when they are created.

## VBA Example



```
Sub SelectAllSectionCuts()
'dimension variables
Dim SapObject As Sap2000v16.SapObject
Dim SapModel As cSapModel
Dim ret As Long
 
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
 
'Deselect section cut for output
ret = SapModel.Results.Setup.SetSectionCutSelectedForOutput("SCut1", False)
 
'Select section cut for output
ret = SapModel.Results.Setup.SetSectionCutSelectedForOutput("SCut1", True)
 
'close Sap2000
SapObject.ApplicationExit False
Set SapModel = Nothing
Set SapObject = Nothing
 
End Sub
 
```



## Release Notes

Initial release in version 16.0.2.

## See Also

[GetSectionCutSelectedForOutput](GetSectionCutSelectedForOutput.md)

[SelectAllSectionCutsForOutput](SelectAllSectionCutsForOutput.md)
