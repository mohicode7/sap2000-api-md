# SetAutoMesh (Note:  Newer function available)

## Syntax

SapObject.SapModel.AreaObj.SetAutoMesh

## VB6 Procedure

Function SetAutoMesh(ByVal Name As String, ByVal MeshType As Long, Optional ByVal n1 As Long = 2, Optional ByVal n2 As Long = 2, Optional ByVal MaxSize1 As Double = 0, Optional ByVal MaxSize2 As Double = 0, Optional ByVal PointOnEdgeFromLine As Boolean = False, Optional ByVal PointOnEdgeFromPoint As Boolean = False, Optional ByVal ExtendCookieCutLines As Boolean = False, Optional ByVal Rotation As Double = 0, Optional ByVal MaxSizeGeneral As Double = 0, Optional ByVal LocalAxesOnEdge As Boolean = False, Optional ByVal LocalAxesOnFace As Boolean = False, Optional ByVal RestraintsOnEdge As Boolean = False, Optional ByVal RestraintsOnFace As Boolean = False, Optional ByVal Group As String = "ALL", Optional ByVal SubMesh As Boolean = False, Optional ByVal SubMeshSize As Double = 0, Optional ByVal ItemType As eItemType = Object) As Long

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

          ret = SapModel.AreaObj.SetAutoMesh("ALL", 1, 3, 3, , , , , , , , , , , , , , , Group)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.00.

This function is obsolete and was superseded by [SetAutoMesh\_1](../Object_Model/Area_Object/SetAutoMesh__1%7BArea_Object%7D.md) in release v27.0.0. This function is maintained for backwards compatibility.

## See Also

[GetAutoMesh](GetAutoMesh_%7BArea_Object%7D.md)
