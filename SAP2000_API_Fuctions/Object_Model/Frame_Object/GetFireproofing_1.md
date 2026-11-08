# GetFireproofing\_1

## Syntax

SapObject.SapModel.FrameObj.GetFireProofing\_1

## VB6 Procedure

Function GetFireProofing\_1(ByVal Name As String, ByRef MyType As Long, ByRef Thickness As Double, ByRef Perimeter As Double, ByRef Density As Double, ByRef tf As Boolean, ByRef IncludeInSelfWeight As Boolean, ByRef IncludeInGravityLoads As Boolean, ByRef IncludeInThisLoadPattern As String) As Long

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

## **Remarks**

This function gets the fireproofing assignment to an existing frame object.

The function returns zero if the fireproofing assignment is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFireproofing()
'dimension variables
Dim SapObject as cOAPI
Dim SapModel As cSapModel
Dim ret As Long
Dim MyType As Long
Dim Thickness As Double
Dim Perimeter As Double
Dim Density As Double
Dim tf As Boolean
Dim IncludeInSelfWeight As Boolean
Dim IncludeInGravityLoads As Boolean
Dim IncludeInThisLoadPattern As String
 
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
ret = SapModel.LoadPatterns.Add("Fireproofing", LTYPE_SUPERDEAD)
 
'assign fireproofing, use in self weight and in load pattern
ret = SapModel.FrameObj.SetFireproofing_1("ALL", 1, 2, 0, 8.68E-06, False, True, False, "Fireproofing", Group)
 
'get fireproofing data for frame object 3
ret = SapModel. FrameObj.GetFireproofing_1("3", MyType, Thickness, Perimeter, Density, tf, IncludeInSelfWeight, IncludeInGravityLoads, IncludeInThisLoadPattern)
 
'close Sap2000
SapObject.ApplicationExit False
Set SapModel = Nothing
Set SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 16.0.0.  This function supersedes obsolete function GetFireProofing

## See Also

[SetFireProofing\_1](SetFireproofing_1.md)

[DeleteFireProofing](DeleteFireproofing.md)
