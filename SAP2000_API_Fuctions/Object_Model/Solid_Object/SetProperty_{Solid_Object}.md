# SetProperty

## Syntax

SapObject.SapModel.SolidObj.SetProperty

## VB6 Procedure

Function SetProperty(ByVal name As String, ByVal PropName As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

PropName

This is the name of a solid property to be assigned to the specified solid object(s).

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the solid object specified by the Name item.

If this item is Group, the assignment is made to all solid objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected solid objects, and the Name item is ignored.

## Remarks

This function assigns a solid property to solid objects.

The function returns zero if the property is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSolidObjectProp()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'set solid property

      ret = SapModel.SolidObj.SetProperty("1", "Solid1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetProperty](GetProperty_%7BSolid_Object%7D.md)
