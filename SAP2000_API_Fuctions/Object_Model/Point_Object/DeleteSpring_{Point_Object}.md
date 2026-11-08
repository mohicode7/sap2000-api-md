# DeleteSpring

## Syntax

SapObject.SapModel.PointObj.DeleteSpring

## VB6 Procedure

Function DeleteSpring(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of a point object or a group depending on the value selected for ItemType item.

ItemType

This is one of the following items from the eItemType enumeration.

Object = 0

Group = 1

SelectedObjects = 2

If Object is selected, the Name item refers to a point object. The point spring assignments for that point object are removed.

If Group is selected, the Name item refers to a group. The point spring assignments for all point objects in the group are removed.

If SelectedObjects is selected, the Name item is ignored. The point spring assignments for all selected point objects are removed.

## Remarks

This function deletes all point spring assignments from the specified point object(s).

The function returns zero if the restraint assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeletePointSprings()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim i as Long

      Dim ret As Long

      Dim k() As Double



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



   'add joint spring assignment

      redim k(5)

      For i = 0 to 5

         k(i) = i + 1

      Next i

      ret = SapModel.PointObj.SetSpring("3", k)



   'delete joint spring assignment

      ret = SapModel.PointObj.DeleteSpring("3")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpring](GetSpring_%7BPoint_Object%7D.md)

[GetSpringCoupled](GetSpringCoupled_%7BPoint_Object%7D.md)

[SetSpring](SetSpring.md)

[SetSpringCoupled](SetSpringCoupled.md)
