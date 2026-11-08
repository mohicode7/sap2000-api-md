# SetInsertionPoint

## Syntax

SapObject.SapModel.CableObj.SetInsertionPoint

## VB6 Procedure

Function SetInsertionPoint(ByVal Name As String, ByVal StiffTransform As Boolean, ByRef Offset1() As Double, ByRef Offset2() As Double, Optional ByVal CSys As String = "Local", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

StiffTransform

If this item is True, the cable object stiffness is transformed for cardinal point and joint offsets from the cable section centroid.

Offset1

This is an array of three joint offset distances, in the coordinate directions specified by CSys, at the I-End of the cable object. [L]

Offset1(0) = Offset in the 1-axis or X-axis direction

Offset1(1) = Offset in the 2-axis or Y-axis direction

Offset1(2) = Offset in the 3-axis or Z-axis direction

Offset2

This is an array of three joint offset distances, in the coordinate directions specified by CSys, at the J-End of the cable object. [L]

Offset2(0) = Offset in the 1-axis or X-axis direction

Offset2(1) = Offset in the 2-axis or Y-axis direction

Offset2(2) = Offset in the 3-axis or Z-axis direction

CSys

This is Local or the name of a defined coordinate system. It is the coordinate system in which the Offset1 and Offset2 items are specified.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the cable object specified by the Name item.

If this item is Group, the assignment is made to all cable objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected cable objects, and the Name item is ignored.

## Remarks

This function assigns cable object insertion point data. The assignments include the end joint offsets.

The function returns zero if the insertion point data is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignCableInsertionPoint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim Offset1() As Double

      Dim Offset2() As Double

      Dim Name As String



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



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'assign cable insertion point

      ReDim Offset1(2)

      ReDim Offset2(2)

      For i = 0 To 2

         Offset1(i)=10 + i

         Offset2(i)=20 + i

      Next i

      ret = SapModel.CableObj.SetInsertionPoint(Name, True, Offset1, Offset2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

This function is obsolete as of v19.2.0 when cable insertion point assignments were removed from the program.

## See Also

[GetInsertionPoint](GetInsertionPoint_%7BCable_Object%7D.md)
