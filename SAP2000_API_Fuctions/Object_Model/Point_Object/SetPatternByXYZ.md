# SetPatternByXYZ

## Syntax

SapObject.Sap2000.PointObj.SetPatternByXYZ

## VB6 Procedure

Function SetPatternByXYZ(ByVal Name As String, ByVal PatternName As String, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, Optional ByVal ItemType As eItemType = object, Optional ByVal Restriction As Long = 0, Optional ByVal Replace As Boolean = False) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

PatternName

The name of a defined joint pattern.

a

The value a in the equation shown in the Remarks section. [1/L]

b

The value b in the equation shown in the Remarks section. [1/L]

c

The value c in the equation shown in the Remarks section. [1/L]

d

The value d in the equation shown in the Remarks section. This item is unitless.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the joint pattern assignment is made to the point object specified by the Name item.

If this item is Group, the joint pattern assignment is made to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the joint pattern assignment is made to all selected point objects and the Name item is ignored.

Restriction

This is either 0, 1, or 2.

0 = All values are used

1 = Negative values are set to zero

2 = Positive values are set to zero

Replace

If this item is True, the joint pattern value calculated as shown in the Remarks section replaces any previous joint pattern value for the point object.

If this item is False, the joint pattern value calculated as shown in the Remarks section is added to any previous joint pattern value for the point object and then the Restriction items are checked.

## Remarks

This function sets the joint pattern value for a specified point object and joint pattern.

The joint pattern value is calculated as:

Value = ax + by + cz + d

where a, b, c and d are function input parameters and x, y and z are the coordinates of the considered point object in the present coordinate system

The function returns zero if the pattern value is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetJointPatternByXYZ()

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



   'add joint pattern assignment

      ret = SapModel.PointObj.SetPatternByXYZ("ALL", "Default", 0, 0, 10, 0, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetPatternValue](GetPatternValue_%7BPoint_Object%7D.md)

[SetPatternByPressure](SetPatternByPressure.md)

[DeletePatternValue](DeletePatternValue.md)

[GetPresentCoordSystem](../../General_Functions/GetPresentCoordSystem.md)

[SetPresentCoordSystem](../../General_Functions/SetPresentCoordSystem.md)
