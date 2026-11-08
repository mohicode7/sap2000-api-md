# DeleteModifiers

## Syntax

SapObject.SapModel.CableObj.DeleteModifiers

## VB6 Procedure

Function DeleteModifiers(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the cable modifier assignments are deleted for the cable object specified by the Name item.

If this item is Group, the cable modifier assignments are deleted for all cable objects in the group specified by the Name item.

If this item is SelectedObjects, the cable modifier assignments are deleted for all selected cable objects, and the Name item is ignored.

## Remarks

This function deletes the cable modifier assignments for cable objects.

The function returns zero if the modifier assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteCableModifiers()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim Value() As Double

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



   'assign modifiers

      ReDim Value(2)

      For i = 0 To 2

         Value(i) = 1

      Next i

      Value(0) = 100

      ret = SapModel.CableObj.SetModifiers(Name, Value)



   'delete modifiers

      ret = SapModel.CableObj.DeleteModifiers(Name)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetModifiers](GetModifiers_%7BCable_Object%7D.md)

[SetModifiers](SetModifiers_%7BCable_Object%7D.md)
