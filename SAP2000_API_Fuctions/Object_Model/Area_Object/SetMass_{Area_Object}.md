# SetMass

## Syntax

SapObject.SapModel.AreaObj.SetMass

## VB6 Procedure

Function SetMass(ByVal Name As String, ByVal MassOverL2 As Double, Optional ByVal Replace As Boolean = False, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

MassOverL2

The mass per unit area assigned to the area object. [M/L2]

Replace

If this item is True, all existing mass assignments to the area object are removed before assigning the specified mas. If it is False, the specified mass is added to any existing mass already assigned to the area object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects and the Name item is ignored.

## Remarks

This function assigns mass per unit area to area objects.

The function returns zero if the mass is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjectMass()

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



   'assign area object mass

      ret = SapModel.AreaObj.SetMass("ALL", .0001, False, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetMass](GetMass_%7BArea_Object%7D.md)

[DeleteMass](DeleteMass_%7BArea_Object%7D.md)
