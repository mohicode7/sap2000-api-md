# SetPropertyFD

## Syntax

SapObject.SapModel.LinkObj.SetPropertyFD

## VB6 Procedure

Function SetPropertyFD(ByVal name As String, ByVal PropName As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing link object or group, depending on the value of the ItemType item.

PropName

This is either None or the name of a frequency dependent link property to be assigned to the specified link object(s). None means that no frequency dependent link property is assigned to the link object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the link object specified by the Name item.

If this item is Group, the assignment is made to all link objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected link objects, and the Name item is ignored.

## Remarks

This function assigns a frequency dependent link property to link objects.

The function returns zero if the property is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetLinkObjectFDProp()

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



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 6-012.sdb")



   'set link frequency dependent property

      ret = SapModel.LinkObj.SetPropertyFD("1", "None")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetProperty](SetProperty_%7BLink_Object%7D.md)

[GetProperty](GetProperty_%7BLink_Object%7D.md)

[GetPropertyFD](GetPropertyFD_%7BLink_Object%7D.md)
