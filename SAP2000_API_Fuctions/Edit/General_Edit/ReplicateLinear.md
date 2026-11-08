# ReplicateLinear

## Syntax

SapObject.SapModel.EditGeneral.ReplicateLinear

## VB6 Procedure

Function ReplicateLinear(ByVal dx As Double, ByVal dy
As Double, ByVal dz As Double, ByVal Number As Long, ByRef NumberObjects
As Long, ByRef ObjectName() As String, ByRef ObjectType() As Long, Optional
ByVal Remove As Boolean = False) As Long

## Parameters

dx, dy, dz

These are the x, y and z offsets used, in the present
coordinate system, to replicate the selected objects.

Number

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

This function linearly replicates selected objects.

The function returns zero if the replication is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub LinearReplication()

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



   'linearly replicate

      ret = SapModel.SelectObj.All

      ret = SapModel.EditGeneral.ReplicateLinear(0, 
 288, 0, 1, NumberObjects, ObjectName, ObjectType)



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

[ReplicateRadial](ReplicateRadial.md)

[ReplicateMirror](ReplicateMirror.md)
