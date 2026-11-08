# SetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.AASHTO\_Concrete\_07.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

Item

This is an integer between 1 and 11, inclusive, indicating the overwrite item considered.

1 = Live load reduction factor

2 = Unbraced length ratio, Major

3 = Unbraced length ratio, Minor

4 = Effective length factor, K Major

5 = Effective length factor, K Minor

6 = Moment coefficient, Cm Major

7 = Moment coefficient, Cm Minor

8 = Nonsway moment factor, Db Major

9 = Nonsway moment factor, Db Minor

10 = Sway moment factor, Ds Major

11 = Sway moment factor, Ds Minor

Value

The value of the considered overwrite item.

1 = Live load reduction factor

Value >= 0; 0 means use program determined value.

2 = Unbraced length ratio, Major

Value >= 0; 0 means use program determined value.

3 = Unbraced length ratio, Minor

Value >= 0; 0 means use program determined value.

4 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

5 = Effective length factor, K Minor

Value >= 0; 0 means use program determined value.

6 = Moment coefficient, Cm Major

Value >= 0; 0 means use program determined value.

7 = Moment coefficient, Cm Minor

Value >= 0; 0 means use program determined value.

8 = Nonsway moment factor, Db Major

Value >= 0; 0 means use program determined value.

9 = Nonsway moment factor, Db Minor

Value >= 0; 0 means use program determined value.

10 = Sway moment factor, Ds Major

Value >= 0; 0 means use program determined value.

11 = Sway moment factor, Ds Minor

Value >= 0; 0 means use program determined value.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, the assignment is made to all selected frame objects and the Name item is ignored.

## Remarks

This function sets the value of a concrete design overwrite item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignOverwriteItem AASHTO_Concrete_07 ()

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



   'create new concrete frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("AASHTO Concrete 07")



   'set overwrite item

      ret = SapModel.DesignConcrete.AASHTO_Concrete_07.SetOverwrite("8", 1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.1.0.

## See Also

[GetOverwrite](GetOverwrite_%7BConcrete_AASHTO_07%7D.md)
