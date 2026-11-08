# ReplicateMirror

## Syntax

SapObject.SapModel.EditGeneral.ReplicateMirror

## VB6 Procedure

Function ReplicateMirror(ByVal Plane As Long, ByVal
x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double,
ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As
Double, ByVal z3 As Double, ByRef NumberObjects As Long, ByRef ObjectName()
As String, ByRef ObjectType() As Long, Optional ByVal Remove As Boolean
= False) As Long

## Parameters

Plane

This is 1, 2, 3 or 4, indicating the mirror plane type.

1 = Parallel to Z

2 = Parallel to X

3 = Parallel to Y

4 = 3D plane

x1, y1, z1, x2, y2, z2, x3, y3, z3

These are the coordinates of three points used to define
the mirror plane. [L]

When Plane = 1, x1, y1, x2 and y2 define the intersection
of the mirror plane with the XY plane.

When Plane = 2, y1, z1, y2 and z2 define the intersection
of the mirror plane with the YZ plane.

When Plane = 3, x1, z1, x2 and z2 define the intersection
of the mirror plane with the XZ plane.

When Plane = 4, x1, y1, z1, x2, y2, z2, x3, y3 and z3
define three points that define the mirror plane.

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

This function mirror replicates selected objects.

The function returns zero if the replication is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub MirrorReplication()

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



   'mirror replicate

      ret = SapModel.SelectObj.All

      ret = SapModel.EditGeneral.ReplicateMirror(1, 
 300, 0, 0, 300, 100, 0, 0, 0, 0, NumberObjects, ObjectName, ObjectType)



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

[ReplicateRadial](ReplicateRadial.md)
