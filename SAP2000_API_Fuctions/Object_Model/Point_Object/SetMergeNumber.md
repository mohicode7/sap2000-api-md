# SetMergeNumber

## Syntax

SapObject.SapModel.PointObj.SetMergeNumber

## VB6 Procedure

Function SetMergeNumber(ByVal Name As String, ByVal MergeNumber As Long, Optional ByVal ItemType As eItemType = object) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

MergeNumber

The merge number for the specified point object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the merge number assignment is made to the point object specified by the Name item.

If this item is Group, the merge number assignment is made to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the merge number assignment is made to all selected point objects and the Name item is ignored.

## Remarks

This function assigns a merge number to a point object. By default the merge number for a point is zero. Points with different merge numbers are not automatically merged by the program.

The function returns zero if the merge number is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignPointMergeNumber()

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



   'set merge number

      ret = SapModel.PointObj.SetMergeNumber("3", 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetMergeNumber](GetMergeNumber_%7BPoint_Object%7D.md)
