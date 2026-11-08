# SetDiscretization

## Syntax

SapObject.SapModel.TendonObj.SetDiscretization

## VB6 Procedure

Function SetDiscretization(ByVal Name As String, ByVal Ang As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing tendon object or group, depending on the value of the ItemType item.

Value

The maximum discretization length for the tendon. [L]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the tendon object specified by the Name item.

If this item is Group, the assignment is made to all tendon objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected tendon objects, and the Name item is ignored.

## Remarks

This function assigns a maximum discretization length to tendon objects.

The function returns zero if the discretization length is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignTendonMaxDiscretizationLength()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'assign tendon maximum discretization length

      ret = SapModel.TendonObj.SetDiscretization(Name, 12)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetDiscretization](GetDiscretization_%7BTendon_Object%7D.md)
