# SetAutoMesh\_1

## Syntax

SapObject.SapModel.AreaObj.SetAutoMesh\_1

## VB6 Procedure

Function GetAutoMesh\_1(ByVal Name As String, ByRef MeshType As Long, ByRef n1 As Long, ByRef n2 As Long, ByRef MaxSize1 As Double, ByRef MaxSize2 As Double, ByRef PointOnEdgeFromLine As Boolean, ByRef PointOnEdgeFromPoint As Boolean, ByRef ExtendCookieCutLines As Boolean, ByRef Rotation As Double, ByRef MaxSizeGeneral As Double, ByRef LocalAxesOnEdge As Boolean, ByRef LocalAxesOnFace As Boolean, ByRef RestraintsOnEdge As Boolean, ByRef RestraintsOnFace As Boolean, ByRef Group As String, ByRef SubMesh As Boolean, ByRef SubMeshSize As Double, ByRef UseUserMesh As Boolean, ByRef QuadsOnlyGeneral As Boolean) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

MeshType

This item is 0, 1, 2, 3, 4, 5 or 6, indicating the automatic mesh type for the area object.

0 = No automatic meshing

1 = Mesh area into a specified number of objects

2 = Mesh area into objects of a specified maximum size

3 = Mesh area based on points on area edges

4 = Cookie cut mesh area based on lines intersecting edges

5 = Cookie cut mesh area based on points

6 = Mesh area using General Divide Tool

Mesh options 1, 2 and 3 apply to quadrilaterals and triangles only.

n1

This item applies when MeshType = 1. It is the number of objects created along the edge of the meshed area object that runs from point 1 to point 2.

n2

This item applies when MeshType = 1. It is the number of objects created along the edge of the meshed area object that runs from point 1 to point 3.

MaxSize1

This item applies when MeshType = 2. It is the maximum size of objects created along the edge of the meshed area object that runs from point 1 to point 2. [L]

If this item is input as 0, the default value is used. The default value is 48 inches if the database units are English or 120 centimeters if the database units are metric.

MaxSize2

This item applies when MeshType = 2. It is the maximum size of objects created along the edge of the meshed area object that runs from point 1 to point 3. [L]

If this item is input as 0, the default value is used. The default value is 48 inches if the database units are English or 120 centimeters if the database units are metric.

PointOnEdgeFromLine

This item applies when MeshType = 3. If it is True, points on the area object edges are determined from intersections of straight line objects included in the group specified by the Group item with the area object edges.

PointOnEdgeFromPoint

This item applies when MeshType = 3. If it is True, points on the area object edges are determined from point objects included in the group specified by the Group item that lie on the area object edges.

ExtendCookieCutLines

This item applies when MeshType = 4. MeshType = 4 provides cookie cut meshing based on straight line objects included in the group specified by the Group item that intersect the area object edges. If the ExtendCookieCutLines item is True, all straight line objects included in the group specified by the Group item are extended to intersect the area object edges for the purpose of meshing the area object.

Rotation

This item applies when MeshType = 5. MeshType = 5 provides cookie cut meshing based on two perpendicular lines passing through point objects included in the group specified by the Group item. By default these lines align with the area object local 1 and 2 axes. The Rotation item is an angle in degrees that the meshing lines are rotated from their default orientation. [deg]

MaxSizeGeneral

This item applies when MeshType = 6. It is the maximum size of objects created by the General Divide Tool.

If this item is input as 0, the default value is used. The default value is 48 inches if the database units are English or 120 centimeters if the database units are metric.

LocalAxesOnEdge

If this item is True, and if both points along an edge of the original area object have the same local axes, then the program makes the local axes for added points along the edge the same as the edge end points.

LocalAxesOnFace

If this item is True, and if all points around the perimeter of the original area object have the same local axes, the program makes the local axes for all added points the same as the perimeter points.

RestraintsOnEdge

If this item is True, and if both points along an edge of the original area object have the same restraint/constraint, then, if the added point and the adjacent corner points have the same local axes definition, the program includes the restraint/constraint for added points along the edge.

RestraintsOnFace

If this item is True, and if all points around the perimeter of the original area object have the same restraint/constraint, then, if an added point and the perimeter points have the same local axes definition, the program includes the restraint/constraint for the added point.

Group

The name of a defined group. Some of the meshing options make use of point and line objects included in this group.

SubMesh

If this item is True, after initial meshing, the program further meshes any area objects that have an edge longer than the length specified by the SubMeshSize item.

SubMeshSize

This item applies when the SubMesh item is True. It is the maximum size of area objects to remain when the auto meshing is complete. [L]

If this item is input as 0, the default value is used. The default value is 12 inches if the database units are English or 30 centimeters if the database units are metric.

UseUserMesh

If this item is True, the user mesh, if assigned, will be used for analysis, otherwise the auto mesh option will be used.

QuadsOnlyGeneral

If this item is True, the general mesh will be generated with quadrilateral areas only, otherwise the mesh may include triangles as necessary. This only applies when MeshType = 6.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects, and the Name item is ignored.

## Remarks

This function makes automatic meshing assignments to area objects.

The function returns zero if the meshing options are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjAutoMesh()

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

          ret = SapModel.File.NewWall(2, 48, 2, 48)



       'assign auto mesh options

          ret = SapModel.AreaObj.SetAutoMesh_1("ALL", 1, 3, 3, , , , , , , , , , , , , , , True, False, Group)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 27.0.0

This function supersedes [SetAutoMesh {Area Object}](../../Obsolete_Functions/SetAutoMesh_%7BArea_Object%7D.md)

## See Also

[GetAutoMesh\_1 {Area Object}](GetAutoMesh_1_%7BArea_Object%7D.md)
