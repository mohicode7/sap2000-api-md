# SetMatTemp

## Syntax

SapObject.SapModel.SolidObj.SetMatTemp

## VB6 Procedure

Function SetMatTemp(ByVal Name As String, ByVal Temp As Double, Optional ByVal PatternName As String = "", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

Temp

This is the material temperature value assigned to the solid object. [T]

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the material temperature for the solid object is uniform over the object at the value specified by Temp.

If PatternName is the name of a defined joint pattern, the material temperature for the solid object may vary. The material temperature at each corner point of the solid object is equal to the specified temperature multiplied by the pattern value at the associated point object. The material temperature at other points in the solid object is calculated by interpolation from the corner points.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the solid object specified by the Name item.

If this item is Group, the assignment is made to all solid objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected solid objects, and the Name item is ignored.

## Remarks

This function assigns material temperatures to solid objects.

The function returns zero if the material temperatures are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSolidObjectMatTemp()

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



   'assign material temperature

      ret = SapModel.SolidObj.SetMatTemp("ALL", 50, , Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetMatTemp](GetMatTemp_%7BSolid_Object%7D.md)
