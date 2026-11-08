# GetAutoMesh

## Syntax

SapObject.SapModel.FrameObj.GetAutoMesh

## VB6 Procedure

Function GetAutoMesh(ByVal Name As String, ByRef AutoMesh As Boolean, ByRef AutoMeshAtPoints As Boolean, ByRef AutoMeshAtLines As Boolean, ByRef NumSegs As Long, ByRef AutoMeshMaxLength As Double) As Long

## Parameters

Name

The name of an existing frame object.

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

## Remarks

This function retrieves the automatic meshing assignments to frame objects.

The function returns zero if the meshing assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameAutoMesh()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim AutoMesh As Boolean

      Dim AutoMeshAtPoints As Boolean

      Dim AutoMeshAtLines As Boolean

      Dim NumSegs As Long

      Dim AutoMeshMaxLength As Double



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



   'get automesh assignment

      ret = SapModel.FrameObj.GetAutoMesh("3", AutoMesh, AutoMeshAtPoints, AutoMeshAtLines,NumSegs, AutoMeshMaxLength)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetAutoMesh](SetAutoMesh_%7BFrame_Object%7D.md)
