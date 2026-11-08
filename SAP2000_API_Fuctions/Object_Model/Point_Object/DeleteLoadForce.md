# DeleteLoadForce

## Syntax

SapObject.SapModel.PointObj.DeleteLoadForce

## VB6 Procedure

Function DeleteLoadForce(ByVal Name As String, ByVal LoadPat As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of a point object or a group depending on the value selected for ItemType item.

LoadPat

The name of a defined load pattern.

ItemType

This is one of the following items from the eItemType enumeration.

Object = 0

Group = 1

SelectedObjects = 2

If Object is selected, the Name item refers to a point object. The point load assignments, for the specified load pattern, made to that point object, are removed.

If Group is selected, the Name item refers to a group. The point load assignments, for the specified load pattern, made to all point objects in the group, are removed.

If SelectedObjects is selected, the Name item is ignored. The point load assignments, for the specified load pattern, made to all selected point objects, are removed.

## Remarks

This function deletes all point load assignments, for the specified load pattern, from the specified point object(s).

The function returns zero if the load assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeletePointForceLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add point load

      Redim Value(5)

      Value(0) = 10

      ret = SapModel.PointObj.SetLoadForce("1", "DEAD", Value)



   'delete point load

      ret = SapModel.PointObj.DeleteLoadForce("1", "DEAD")



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

[GetLoadForce](GetLoadForce_%7BPoint_Object%7D.md)

[SetLoadForce](SetLoadForce.md)
