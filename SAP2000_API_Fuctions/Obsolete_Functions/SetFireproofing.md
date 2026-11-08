# SetFireproofing

## Syntax

SapObject.SapModel.FrameObj.SetFireproofing

## VB6 Procedure

Function SetFireproofing(ByVal Name As String, ByVal MyType As Long, ByVal Thickness As Double, ByVal Perimeter As Double, ByVal Density As Double, ByVal tf As Boolean, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

MyType

This is 1, 2 or 3, indicating the type of fireproofing assigned.

1 = Sprayed on - program calculate section perimeter

2 = Sprayed on - user provides section perimeter

3 = Concrete encased

Thickness

When MyType = 1 or MyType = 2, this is the thickness of the sprayed on fireproofing. When MyType = 3, this is the concrete cover dimension. [L]

Perimeter

This item applies only when MyType = 2. It is the length of fireproofing applied, measured around the perimeter of the frame object cross-section. [L]

Density

This is the weight per unit volume of the fireproofing material. [F/L3]

tf

This item applies only when MyType = 1 or MyType = 3. If this item is True, the fireproofing is assumed to be applied to the top flange of the section. If it is False, the program assumes no fireproofing is applied to the section top flange. This flag applies for I, channel and double channel sections.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function assigns fireproofing to frame objects.

The function returns zero if the fireproofing assignments are successfully assigned, otherwise it returns a nonzero value.

The program automatically adds the load \*(weight) calculated for the fireproofing to all load patterns that include self weight.

## VBA Example



```
Sub AssignFireproofing()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'assign fireproofing

      ret = SapModel.FrameObj.SetFireproofing("ALL", 1, 2, 0, 8.68E-06, False, Group)



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

[GetFireproofing](GetFireproofing.md)

[DeleteFireproofing](../Object_Model/Frame_Object/DeleteFireproofing.md)
