# SetFireproofing\_1

## Syntax

SapObject.SapModel.FrameObj.SetFireProofing\_1

## VB6 Procedure

Function SetFireProofing\_1(ByVal Name As String, ByVal MyType As Long, ByVal Thickness As Double, ByVal Perimeter As Double, ByVal Density As Double, ByVal tf As Boolean, ByVal IncludeInSelfWeightAs Boolean, ByVal IncludeInGravityLoads As Boolean, Optional ByVal IncludeInThisLoadPattern As String = "None", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object.

MyType

This is 1, 2 or 3, indicating the type of fireproofing assigned.

1 = Sprayed on - program calculate section perimeter

2 = Sprayed on - user provides section perimeter

3 = Concrete encased

Thickness

When MyType = 1 or MyType = 2 this is the thickness of the sprayed on fireproofing. When MyType = 3 this is the concrete cover dimension. [L]

Perimeter

This item applies only when MyType = 2. It is the length of fireproofing applied measured around the perimeter of the frame object cross-section. [L]

Density

This is the weight per unit volume of the fireproofing material. [F/L3]

tf

This item  applies only when MyType = 1 or MyType = 3. If this item is True, the fireproofing is assumed to be applied to the top flange of the section. If it is False, the program assumes no fireproofing is applied to the section top flange. This flag applies for I, channel and double channel sections.

IncludeInSelfWeight

If this item is True the fireproofing is included in the structure self weight.

IncludeInGravityLoads

If this item is True the fireproofing is included gravity loads applied in the X, Y and Z directions.

IncludeInThisLoadPattern

This item is either None or the name of an existing load pattern. If it is the name of a load pattern then the weight of the fireproofing is applied as a distributed load in the global Z direction in the load pattern.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## **Remarks**

This function sets the fireproofing assignments to existing frame objects.

The function returns zero if the fireproofing assignments are successfully set, otherwise it returns a nonzero value.

## VBA Example

Sub SetFireproofing()

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

ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)

'add new load pattern

ret = SapModel.LoadPatterns.Add("Fireproofing", LTYPE\_SUPERDEAD)

'assign fireproofing, use in self weight and in load pattern

ret = SapModel.FrameObj.SetFireproofing\_1("ALL", 1, 2, 0, 8.68E-06, False, True, False, "Fireproofing", Group)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.  This function supersedes obsolete function SetFireProofing

## See Also

[GetFireProofing\_1](GetFireproofing_1.md)

[DeleteFireProofing](DeleteFireproofing.md)
