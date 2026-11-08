# SetAutoMesh

## Syntax

SapObject.SapModel.FrameObj.SetAutoMesh

## VB6 Procedure

Function SetAutoMesh(ByVal Name As String, ByVal AutoMesh As Boolean, ByVal AutoMeshAtPoints As Boolean, ByVal AutoMeshAtLines As Boolean, ByVal NumSegs As Long, ByVal AutoMeshMaxLength As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

AutoMesh

This item is True if the frame object is to be automatically meshed by the program when the analysis model is created.

AutoMeshAtPoints

This item is applicable only when the AutoMesh item is True. If this item is True, the frame object is automatically meshed at intermediate joints along its length.

AutoMeshAtLines

This item is applicable only when the AutoMesh item is True. If this item is True, the frame object is automatically meshed at intersections with other frames, area object edges and solid object edges.

NumSegs

This item is applicable only when the AutoMesh item is True. It is the minimum number of elements into which the frame object is automatically meshed. If this item is zero, the number of elements is not checked when the automatic meshing is done.

AutoMeshMaxLength

This item is applicable only when the AutoMesh item is True. It is the maximum length of auto meshed frame elements. If this item is zero, the element length is not checked when the automatic meshing is done. [L]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function makes automatic meshing assignments to frame objects.

The function returns zero if the meshing options are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameAutoMesh()

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



   'assign automesh options

      ret = SapModel.FrameObj.SetAutoMesh("ALL", True, True, True, 0, 0, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetAutoMesh](GetAutoMesh_%7BFrame_Object%7D.md)
