# DeleteCut

## Syntax

SapObject.SapModel.SectCut.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of the section cut to be deleted.

## Remarks

This function deletes an existing section cut.

The function returns zero if the section cut is successfully
deleted, otherwise it returns a nonzero value.

## VBA Example

Sub
DeleteSectionCut()  
  
'dimension
variables  
      Dim SapObject as cOAPI  
      Dim SapModel As cSapModel  
      Dim ret As Long  
      Dim X(3) As Double  
      Dim Y(3) As Double  
      Dim Z(3) As Double  
  
'create
Sap2000 object  
      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")  
  
'start
Sap2000 application  
      SapObject.ApplicationStart  
  
'create
SapModel object  
      Set SapModel = SapObject.SapModel  
  
'initialize
model  
      ret
= SapModel.InitializeNewModel  
  
'create
model from template  
      ret = SapModel.File.New2DFrame(PortalFrame,
2, 144, 2, 288)

'define
new group  
 ret = SapModel.GroupDef.SetGroup("Group1")

'add objects to
group  
 ret = SapModel.FrameObj.SetGroupAssign("1",
"Group1")  
 ret = SapModel.FrameObj.SetGroupAssign("2", "Group1")  
 ret = SapModel.FrameObj.SetGroupAssign("3", "Group1")

'define section
cut  
 X(0) = -300  
 X(1) = -270  
 X(2) = -270  
 X(3) = -300  
 Y(0) = 10  
 Y(1) = 10  
 Y(2) = -10  
 Y(3) = -10  
 Z(0) = 10  
 Z(1) = 10  
 Z(2) = 10  
 Z(3) = 10  
 ret = SapModel.SectCut.SetByQuad("SCut1", "Group1",
1, X, Y, Z)

'delete section
cut  
 ret = SapModel.SectCut.Delete("SCut1")

'close Sap2000  
 SapObject.ApplicationExit False  
 Set SapModel = Nothing  
 Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 23.4.0.

## See Also

[SetByQuad](SetByQuad.md)
