# SetMaterialOverwrite

## Syntax

SapObject.SapModel.FrameObj.SetMaterialOverwrite

## VB6 Procedure

Function SetMaterialOverwrite(ByVal Name As String, ByVal PropName As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

PropName

This is None or a blank string, indicating that any existing material overwrites assigned to the specified frame objects are to be removed, or it is the name of an existing material property.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function sets the material overwrite assignment for frame objects.

The function returns zero if the material overwrite assignment is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameMaterialOverwrite()

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



   'assign material overwrite

      ret = SapModel.FrameObj.SetMaterialOverwrite("3", "4000Psi")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetMaterialOverwrite](GetMaterialOverwrite_%7BFrame_Object%7D.md)
