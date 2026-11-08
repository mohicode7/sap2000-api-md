# SetEndSkew

## Syntax

SapObject.SapModel.FrameObj.SetEndSkew

## VB6 Procedure

Function SetEndSkew(ByVal Name As String, ByVal SkewI As Double, ByVal SkewJ As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

SkewI

The angle in degrees measured counter clockwise from the positive local 3-axis to a line parallel to the I-End of the frame object (-90 < SkewI < 90). [deg]

SkewJ

The angle in degrees measured counter clockwise from the positive local 3-axis to a line parallel to the J-End of the frame object (-90 < SkewJ < 90). [deg]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function assigns frame object end skew data. End skew data is used in the program to plot the extruded view of bridge objects that have been updated as spine models only.

The function returns zero if the end skew data is successfully assigned, otherwise it returns a nonzero value.

End skew assignments are applicable only to straight frame objects.

## VBA Example



```
Sub AssignFrameEndSkewData()

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



   'assign frame end skew data

      ret = SapModel.FrameObj.SetEndSkew("15", 10, 20)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetEndSkew](GetEndSkew.md)
