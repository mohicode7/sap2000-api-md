# SetMass

## Syntax

SapObject.SapModel.CableObj.SetMass

## VB6 Procedure

Function SetMass(ByVal Name As String, ByVal MassOverL As Double, Optional ByVal Replace As Boolean = False, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

MassOverL

The mass per unit length assigned to the cable object. [M/L]

Replace

If this item is True, all existing mass assignments to the cable object are removed before assigning the specified mas. If it is False, the specified mass is added to any mass already assigned to the cable object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the cable object specified by the Name item.

If this item is Group, the assignment is made to all cable objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected cable objects, and the Name item is ignored.

## Remarks

This function assigns mass per unit length to cable objects.

The function returns zero if the mass is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignCableMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



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



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'assign cable mass

      ret = SapModel.CableObj.SetMass("ALL", .0001, False, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetMass](GetMass_%7BCable_Object%7D.md)

[DeleteMass](DeleteMass_%7BCable_Object%7D.md)
