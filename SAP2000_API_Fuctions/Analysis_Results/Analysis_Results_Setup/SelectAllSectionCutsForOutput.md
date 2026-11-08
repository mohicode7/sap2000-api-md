# SelectAllSectionCutsForOutput

## Syntax

SapObject.SapModel.Results.Setup.SelectAllSectionCutsForOutput

## VB6 Procedure

Function SelectAllSectionCutsForOutput (Selected as Boolean) AsLong

## Parameters

Selected

This item is True if all section cuts are to be selected for output, or False if no section cuts are to be selected for output.

## Remarks

This function selects or deselects all section cuts for output.

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
 
'Deselect all section cuts for output
ret = SapModel.Results.Setup.SelectAllSectionCutsForOutput(False)
 
'Select all section cuts for output
ret = SapModel.Results.Setup.SelectAllSectionCutsForOutput(True)
 
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

[SetSectionCutSelectedForOutput](SetSectionCutSelectedForOutput.md)















