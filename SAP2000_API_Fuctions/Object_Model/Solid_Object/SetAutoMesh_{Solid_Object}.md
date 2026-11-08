# SetAutoMesh

## Syntax

SapObject.SapModel.SolidObj.SetAutoMesh

## VB6 Procedure

Function SetAutoMesh(ByVal Name As String, ByVal MeshType As Long, Optional ByVal n1 As Long = 2, Optional ByVal n2 As Long = 2, Optional ByVal n3 As Long = 2, Optional ByVal MaxSize1 As Double = 0, Optional ByVal MaxSize2 As Double = 0, Optional ByVal MaxSize3 As Double = 0, Optional ByVal RestraintsOnEdge As Boolean = False, Optional ByVal RestraintsOnFace As Boolean = False) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

MeshType

This item is 0, 1 or 2, indicating the automatic mesh type for the solid object.

0 = No automatic meshing

1 = Mesh solid into a specified number of objects

2 = Mesh solid into objects of a specified maximum size

n1

This item applies when MeshType = 1. It is the number of objects created along the edge of the meshed solid object that runs from point 1 to point 2.

n2

This item applies when MeshType = 1. It is the number of objects created along the edge of the meshed solid object that runs from point 1 to point 3.

n3

This item applies when MeshType = 1. It is the number of objects created along the edge of the meshed solid object that runs from point 1 to point 5.

MaxSize1

This item applies when MeshType = 2. It is the maximum size of objects created along the edge of the meshed solid object that runs from point 1 to point 2. [L]

If this item is input as 0, the default value is used. The default value is 48 inches if the database units are English or 120 centimeters if the database units are metric.

MaxSize2

This item applies when MeshType = 2. It is the maximum size of objects created along the edge of the meshed solid object that runs from point 1 to point 3. [L]

If this item is input as 0, the default value is used. The default value is 48 inches if the database units are English or 120 centimeters if the database units are metric.

MaxSize3

This item applies when MeshType = 2. It is the maximum size of objects created along the edge of the meshed solid object that runs from point 1 to point 5. [L]

If this item is input as 0, the default value is used. The default value is 48 inches if the database units are English or 120 centimeters if the database units are metric.

RestraintsOnEdge

If this item is True, and if both points along an edge of the original solid object have the same restraint/constraint, then, if the an added point on that edge and the original corner points have the same local axes definition, the program assigns the restraint/constraint to the added point.

RestraintsOnFace

If this item is True, and if all corner points on an solid object face have the same restraint/constraint, then, if an added point on that face and the original corner points for the face have the same local axes definition, the program assigns the restraint/constraint to the added point.

## Remarks

This function makes automatic meshing assignments to solid objects.

The function returns zero if the meshing options are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSolidObjAutoMesh()

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



   'assign auto mesh options

      ret = SapModel.SolidObj.SetAutoMesh("ALL", 1, 3, 3,3, , , , , , Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetAutoMesh](GetAutoMesh_%7BSolid_Object%7D.md)
