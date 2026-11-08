# DeleteConstraint

## Syntax

SapObject.SapModel.PointObj.DeleteConstraint

## VB6 Procedure

Function DeleteConstraint(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of a point object or a group depending on the value selected for ItemType item.

ItemType

This is one of the following items from the eItemType enumeration.

Object = 0

Group = 1

SelectedObjects = 2

If Object is selected, the Name item refers to a point object and the constraint assignments to that point object are removed.

If Group is selected, the Name item refers to a group and the constraint assignments to all point objects in the group are removed.

If SelectedObjects is selected, the Name item is ignored and the constraint assignments to all selected point objects are removed.

## Remarks

This function deletes all constraint assignments from the specified point object(s).

The function returns zero if the constraint assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeletePointConstraints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'add constraint definition

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1")



   'make constraint assignment

      ret = SapModel.PointObj.SetConstraint("3", "Diaph1")



   'delete constraint assignment

      ret = SapModel.PointObj.DeleteConstraint("3")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetConstraint](GetConstraint_%7BPoint_Object%7D.md)

[SetConstraint](SetConstraint.md)
