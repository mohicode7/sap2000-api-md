# DeleteSpring

## Syntax

SapObject.SapModel.SolidObj.DeleteSpring

## VB6 Procedure

Function DeleteSpring(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the spring assignments are deleted for the solid object specified by the Name item.

If this item is Group, the spring assignments are deleted for all solid objects in the group specified by the Name item.

If this item is SelectedObjects, the spring assignments are deleted for all selected solid objects, and the Name item is ignored.

## Remarks

This function deletes all spring assignments for the specified solid objects.

The function returns zero if the assignments are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteSolidObjectSprings()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Vec() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'assign springs to solid objects

      ReDim Vec(2)

      ret = SapModel.SolidObj.SetSpring("1", 1, 1, 1, "", 1, 1, 3, True, Vec, 0, False, "Local")

      ret = SapModel.SolidObj.SetSpring("5", 1, 1, 1, "", 1, 1, 3, True, Vec, 0, False, "Local")



   'delete springs

      ret = SapModel.SolidObj.DeleteSpring("1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpring](GetSpring_%7BSolid_Object%7D.md)

[SetSpring](SetSpring_%7BSolid_Object%7D.md)
