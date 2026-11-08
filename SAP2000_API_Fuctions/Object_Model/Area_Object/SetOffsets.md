# SetOffsets

## Syntax

SapObject.SapModel.AreaObj.SetOffsets

## VB6 Procedure

Function SetOffsets(ByVal Name As String, ByVal OffsetType As Long, ByVal OffsetPattern As String, ByVal OffsetPatternSF As Double, ByRef Offset() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

OffsetType

This is 0, 1 or 2, indicating the joint offset type.

0 = No joint offsets

1 = User defined joint offsets specified by joint pattern

2 = User defined joint offsets specified by point

OffsetPattern

This item applies only when OffsetType = 1. It is the name of the defined joint pattern that is used to calculate the joint offsets.

OffsetPatternSF

This item applies only when OffsetType = 1. It is the scale factor applied to the joint pattern when calculating the joint offsets. [L]

Offset

This item applies only when OffsetType = 2. It is an array of joint offsets for each of the points that define the area object. [L]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects, and the Name item is ignored.

## Remarks

This function sets the joint offset assignments for area objects.

The function returns zero if the offsets are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjectJointOffsets()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i as long

      Dim Offset() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'assign joint offsets

      ReDim Offset(3)

      For i = 0 To 3

         Offset(i) = 12

      Next i

      ret = SapModel.AreaObj.SetOffsets("ALL", 2, "", 1, Offset, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetOffsets](GetOffsets_%7BArea_Object%7D.md)
