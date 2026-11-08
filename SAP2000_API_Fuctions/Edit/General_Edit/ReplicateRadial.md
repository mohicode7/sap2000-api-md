# ReplicateRadial

## Syntax

SapObject.SapModel.EditGeneral.ReplicateRadial

## VB6 Procedure

Function ReplicateRadial(ByVal RotateAxis As Long, ByVal
x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double,
ByVal y2 As Double, ByVal z2 As Double, ByVal Number As Long, ByVal Ang
As Double, ByRef NumberObjects As Long, ByRef ObjectName() As String,
ByRef ObjectType() As Long, Optional ByVal Remove As Boolean = False)
As Long

## Parameters

RotateAxis

This is 1, 2, 3 or 4, indicating the rotation axis.

1 = Parallel to X axis

2 = Parallel to Y axis

3 = Parallel to Z axis

4 = 3D line

x1, y1, z1

These are coordinates used to define the rotation axis.
[L]

When RotateAxis = 1, y1 and z1 define the intersection
of the rotation axis with the YZ plane.

When RotateAxis = 2, x1 and z1 define the intersection
of the rotation axis with the XZ plane.

When RotateAxis = 3, x1 and y1 define the intersection
of the rotation axis with the XY plane.

When RotateAxis = 4, x1, y1 and z1 define one point
on the rotation axis.

x2, y2, z2

These are coordinates used to define the rotation axis
when RotateAxis = 4. x2, y2 and z2 define a second point on the rotation
axis. [L]

Number

The increment angle for each replication.

Ang

The number of times the selected objects are to be replicated.

NumberObjects

The number of new objects created by the replication
process.

ObjectName

This is an array of the name of each object created
by the replication process.

ObjectType

This is an array of the type of each object created
by the replication process.

1 = Point object

2 = Frame object

3 = Cable object

4 = Tendon object

5 = Area object

6 = Solid object

7 = Link object

Remove

If this item is True, the originally selected objects
are deleted after the replication is complete.

## Remarks

This function radially replicates selected objects.

The function returns zero if the replication is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub RadialReplication()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberObjects As Long

      Dim ObjectName() As String

      Dim ObjectType() As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'radially replicate

      ret = SapModel.SelectObj.All

      ret = SapModel.EditGeneral.ReplicateRadial(3, 
 -360, 0, 0, 0, 0, 0, 2, 45, NumberObjects, ObjectName, ObjectType)



   'refresh view

      ret = SapModel.View.RefreshView(0, 
 False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Modified optional argument Remove to be ByVal in version
12.0.1.

## See Also

[ReplicateLinear](ReplicateLinear.md)

[ReplicateMirror](ReplicateMirror.md)
