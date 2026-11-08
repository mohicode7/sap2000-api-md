# SetPatternByPressure

## Syntax

SapObject.SapModel.PointObj.SetPatternByPressure

## VB6 Procedure

Function SetPatternByPressure(ByVal Name As String, ByVal PatternName As String, ByVal Z As Double, ByVal w As Double, u As Double, r As Long, Optional ByVal ItemType As eItemType = Object, Optional ByVal Restriction As Long = 0, Optional ByVal Replace As Boolean = False) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

PatternName

The name of a defined joint pattern.

z

The Z coordinate at zero pressure in the present coordinate system. [L]

w

A weight per unit volume. [F/L3]

u

An added uniform force per unit area. [F/L2]

r

This is 0, 1, or 2.

0 = All values are used

1 = Negative values are set to zero

2 = Positive values are set to zero

This restriction applies ***before*** the pattern value has been added to any existing pattern value assigned to the point object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object,  the joint pattern assignment is made to the point object specified by the Name item.

If this item is Group, the joint pattern assignment is made to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the joint pattern assignment is made to all selected point objects and the Name item is ignored.

Restriction

This is 0, 1, or 2.

0 = All values are used

1 = Negative values are set to zero

2 = Positive values are set to zero

This restriction applies ***after*** the pattern value has been added to any existing pattern value assigned to the point object. This restriction applies even if there was no existing joint pattern value on the point object.

Replace

If this item is True, the joint pattern value calculated as shown in the Remarks section replaces any previous joint pattern value for the point object.

If this item is False, the joint pattern value calculated as shown in the Remarks section is added to any previous joint pattern value for the point object and then the Restriction items are checked.

## Remarks

This function sets the joint pattern value for a specified point object and joint pattern.

The joint pattern value is calculated as:

Value = [(z – zpoint) \* w] + u

where z, w and u are described in the Parameters section and zpoint is the Z coordinate of the considered point object in the present coordinate system. All appropriate unit conversions are used to calculate the value in the database units, but thereafter it is assumed to be unitless.

The function returns zero if the pattern value is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetJointPatternByPressure()

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

      ret = SapModel.PointObj.SetPatternByPressure("ALL", "Default", 0, 20, 1, 0, Group)



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

[SetPatternByXYZ](SetPatternByXYZ.md)

[DeletePatternValue](DeletePatternValue.md)

[GetPresentCoordSystem](../../General_Functions/GetPresentCoordSystem.md)

[SetPresentCoordSystem](../../General_Functions/SetPresentCoordSystem.md)
