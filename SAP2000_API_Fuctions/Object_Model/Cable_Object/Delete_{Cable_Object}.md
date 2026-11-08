# Delete

## Syntax

SapObject.SapModel.CableObj.Delete

## VB6 Procedure

Function Delete(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the cable object specified by the Name item is deleted.

If this item is Group, the all cable objects in the group specified by the Name item are deleted.

If this item is SelectedObjects, all selected cable objects are deleted, and the Name item is ignored.

## Remarks

The function deletes cable objects.

The function returns zero if the cable objects are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteCableObj()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name1 As String

      Dim Name2 As String



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



   'add cable objects by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name1)

      ret = SapModel.CableObj.AddByPoint("5", "10", Name2)



   'delete cable object

      ret = SapModel.CableObj.Delete(Name1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByCoord](AddByCoord_%7BCable_Object%7D.md)

[AddByPoint](AddByPoint_%7BCable_Object%7D.md)
