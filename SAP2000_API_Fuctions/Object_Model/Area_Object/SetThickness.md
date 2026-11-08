# SetThickness

## Syntax

SapObject.SapModel.AreaObj.SetThickness

## VB6 Procedure

Function SetThickness(ByVal Name As String, ByVal ThicknessType As Long, ByVal ThicknessPattern As String, ByVal ThicknessPatternSF As Double, ByRef Thickness() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

ThicknessType

This is 0, 1 or 2, indicating the thickness overwrite type.

0 = No thickness overwrites

1 = User defined thickness overwrites specified by joint pattern

2 = User defined thickness overwrites specified by point

ThicknessPattern

This item applies only when ThicknessType = 1. It is the name of the defined joint pattern that is used to calculate the thicknesses.

ThicknessPatternSF

This item applies only when ThicknessType = 1. It is the scale factor applied to the joint pattern when calculating the thicknesses. [L]

Thickness

This item applies only when ThicknessType = 2. It is an array of thicknesses at each of the points that define the area object. [L]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects, and the Name item is ignored.

## Remarks

This function sets the thickness overwrite assignments for area objects.

The function returns zero if the thickness overwrites are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjectThicknessOverwrites()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i as long

      Dim Thickness() As Double



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



   'assign thickness overwrites

      ReDim Thickness(3)

      For i = 0 To 3

         Thickness(i) = 11

      Next i

      ret = SapModel.AreaObj.SetThickness("ALL", 2, "", 1, Thickness, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetThickness](GetThickness_%7BArea_Object%7D.md)
