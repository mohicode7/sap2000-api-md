# SetSpring

## Syntax

SapObject.SapModel.AreaObj.SetSpring

## VB6 Procedure

Function SetSpring(ByVal Name As String, ByVal MyType As Long, ByVal s As Double, ByVal SimpleSpringType As Long, ByVal LinkProp As String, ByVal Face as Long, ByVal SpringLocalOneType As Long, ByVal Dir As Long, ByVal Outward As Boolean, ByRef Vec() As Double, ByVal Ang As Double, ByVal Replace As Boolean, Optional ByVal CSys As String = "Local", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

MyType

This is either 1 or 2, indicating the spring property type.

1 = Simple spring

2 = Link property

s

The simple spring stiffness per unit area of the specified area object face. This item applies only when MyType = 1. [F/L3]

SimpleSpringType

This is 1, 2 or 3, indicating the simple spring type. This item applies only when MyType = 1.

1 = Spring resists tension and compression

2 = Spring resists compression only

3 = Spring resists tension only

LinkProp

The name of the link property assigned to the spring. This item applies only when MyType = 2.

Face

This is -1, -2 or a nonzero, positive integer indicating the area object face to which the specified spring assignment applies.

-1 = Bottom face

-2 = Top face

>0 = Edge face

Note that edge face n is from area object point n to area object point n + 1. For example, edge face 2 is from area object point 2 to area object point 3.

SpringLocalOneType

This is 1, 2 or 3, indicating the method used to specify the spring positive local 1-axis orientation.

1 = Parallel to area object local axis

2 = Normal to specified area object face

3 = User specified direction vector

Dir

This is 1, 2, 3, -1, -2 or -3, indicating the area object local axis that corresponds to the positive local 1-axis of the spring. This item applies only when SpringLocalOneType = 1.

Outward

This item is True if the spring positive local 1 axis is outward from the specified area object face. This item applies only when SpringLocalOneType = 2.

Vec

This is an array of three values that define the direction vector of the spring positive local 1-axis. The direction vector is in the coordinate system specified by the CSys item. This item applies only when SpringLocalOneType = 3.

Ang

This is the angle that the link local 2-axis is rotated from its default orientation. This item applies only when MyType = 2. [deg]

Replace

If this item is True, all existing spring assignments to the area object are removed before assigning the specified spring. If it is False, the specified spring is added to any existing springs already assigned to the area object.

CSys

This is Local (meaning the area object local coordinate system) or the name of a defined coordinate system. This item is the coordinate system in which the user specified direction vector, Vec, is specified. This item applies only when SpringLocalOneType = 3.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects, and the Name item is ignored.

## Remarks

This function makes spring assignments to area objects. The springs are assigned to a specified area object face.

The function returns zero if the assignments are successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjectSprings()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'assign springs to area objects

      ReDim Vec(2)

      ret = SapModel.AreaObj.SetSpring("ALL", 1, 1, 1, "", -1, 1, 3, True, Vec, 0, False, "Local", Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpring](GetSpring_%7BArea_Object%7D.md)

[DeleteSpring](DeleteSpring_%7BArea_Object%7D.md)
