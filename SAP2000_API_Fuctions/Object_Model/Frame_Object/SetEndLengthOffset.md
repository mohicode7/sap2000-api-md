# SetEndLengthOffset

## Syntax

SapObject.SapModel.FrameObj.SetEndLengthOffset

## VB6 Procedure

Function SetEndLengthOffset(ByVal Name As String, ByVal AutoOffset As Boolean, ByVal Length1 As Double, ByVal Length2 As Double, ByVal rz As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

AutoOffset

If this item is True, the end length offsets are automatically determined by the program from object connectivity, and the Length1, Length2 and rz items are ignored.

Length1

The offset length along the 1-axis of the frame object at the I-End of the frame object. [L]

Length2

The offset along the 1-axis of the frame object at the J-End of the frame object. [L]

rz

The rigid zone factor.  This is the fraction of the end offset length assumed to be rigid for bending and shear deformations.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function assigns frame object end offsets along the 1-axis of the object.

The function returns zero if the offsets are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameEndOffsets()

   'dimension variables

      Dim SapObject as cOAPI

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'assign offsets

      ret = SapModel.FrameObj.SetEndLengthOffset("15", False, 12, 12, 0.5)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetEndLengthOffset](GetEndLengthOffset_%7BFrame_Object%7D.md)
