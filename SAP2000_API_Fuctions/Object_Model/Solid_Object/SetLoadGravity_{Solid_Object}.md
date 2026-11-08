# SetLoadGravity

## Syntax

SapObject.SapModel.SolidObj.SetLoadGravity

## VB6 Procedure

Function SetLoadGravity(ByVal Name As String, ByVal LoadPat As String, ByVal x As Double, ByVal y As Double, ByVal z As Double, Optional ByVal Replace As Boolean = True, Optional ByVal CSys As String = "Global", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

x, y, z

These are the gravity load multipliers in the x, y and z directions of the specified coordinate system.

Replace

If this item is True, all previous gravity loads, if any, assigned to the specified solid object(s), in the specified load pattern, are deleted before making the new assignment.

CSys

The coordinate system in which the x, y and z multipliers are specified.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the solid object specified by the Name item.

If this item is Group, the assignment is made to all solid objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected solid objects, and the Name item is ignored.

## Remarks

This function assigns gravity load multipliers to solid objects.

The function returns zero if the loads are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSolidObjectGravityLoad()

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



   'assign solid object gravity loads

      ret = SapModel.SolidObj.SetLoadGravity("ALL", "DEAD", 0, 0, -1, , , Group)



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

[GetLoadGravity](GetLoadGravity_%7BSolid_Object%7D.md)

[DeleteLoadGravity](DeleteLoadGravity_%7BSolid_Object%7D.md)
