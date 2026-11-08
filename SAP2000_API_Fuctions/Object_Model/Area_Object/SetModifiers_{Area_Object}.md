# SetModifiers

## Syntax

SapObject.SapModel.AreaObj.SetModifiers

## VB6 Procedure

Function SetModifiers(ByVal Name As String, ByRef Value() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

Value

This is an array of ten unitless modifiers.

Value(0) = Membrane f11 modifier

Value(1) = Membrane f22 modifier

Value(2) = Membrane f12 modifier

Value(3) = Bending m11 modifier

Value(4) = Bending m22 modifier

Value(5) = Bending m12 modifier

Value(6) = Shear v13 modifier

Value(7) = Shear v23 modifier

Value(8) = Mass modifier

Value(9) = Weight modifier

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects, and the Name item is ignored.

## Remarks

This function sets the modifier assignment for area objects. The default value for all modifiers is one.

The function returns zero if the modifier assignments are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaModifiers()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

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



   'assign modifiers

      ReDim Value(9)

      For i = 0 To 9

         Value(i) = 1

      Next i

      Value(0) = 0.01

      ret = SapModel.AreaObj.SetModifiers("ALL", Value, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetModifiers](GetModifiers_%7BArea_Object%7D.md)

[DeleteModifiers](DeleteModifiers_%7BArea_Object%7D.md)
