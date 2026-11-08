# GetLoadUniform

## Syntax

SapObject.SapModel.AreaObj.GetLoadUniform

## VB6 Procedure

Function GetLoadUniform(ByVal Name As String, ByRef NumberItems As Long, ByRef AreaName() As String, ByRef LoadPat() As String, ByRef CSys() As String, ByRef Dir() As Long, ByRef Value() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

NumberItems

The total number of uniform loads retrieved for the specified area objects.

AreaName

This is an array that includes the name of the area object associated with each uniform load.

LoadPat

This is an array that includes the name of the coordinate system in which the uniform load is specified.

CSys

This is an array that includes the name of the coordinate system associated with each uniform load.

Dir

This is an integer between 1 and 11, indicating the direction of the load.

1 = Local 1 axis (only applies when CSys is Local)

2 = Local 2 axis (only applies when CSys is Local)

3 = Local 3 axis (only applies when CSys is Local)

4 = X direction (does not apply when CSys is Local)

5 = Y direction (does not apply when CSys is Local)

6 = Z direction (does not apply when CSys is Local)

7 = Projected X direction (does not apply when CSys is Local)

8 = Projected Y direction (does not apply when CSys is Local)

9 = Projected Z direction (does not apply when CSys is Local)

10 = Gravity direction (only applies when CSys is Global)

11 = Projected Gravity direction (only applies when CSys is Global)

The positive gravity direction (see Dir = 10 and 11) is in the negative Global Z direction.

Value

The uniform load value. [F/L2]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the area object specified by the Name item.

If this item is Group, the assignments are retrieved for all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected area objects, and the Name item is ignored.

## Remarks

This function retrieves the uniform load assignments to area objects.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectUniformLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim AreaName() As String

      Dim LoadPat() As String

      Dim CSys() As String

      Dim Dir() As Long

      Dim Value() As Double



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



   'assign area object uniform loads

      ret = SapModel.AreaObj.SetLoadUniform("ALL", "DEAD", -0.01, 2, False, "Local", Group)



   'get area object uniform load

      ret = SapModel.AreaObj.GetLoadUniform("3", NumberItems, AreaName, LoadPat, CSys, Dir, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SetLoadUniform](SetLoadUniform.md)

[DeleteLoadUniform](DeleteLoadUniform.md)
