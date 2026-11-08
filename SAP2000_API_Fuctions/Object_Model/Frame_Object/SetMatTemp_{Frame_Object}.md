# SetMatTemp

## Syntax

SapObject.SapModel.FrameObj.SetMatTemp

## VB6 Procedure

Function SetMatTemp(ByVal Name As String, ByVal Temp As Double, Optional ByVal PatternName As String = "", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

Temp

This is the material temperature value assigned to the frame object. [T]

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the material temperature for the frame object is uniform along the object at the value specified by Temp.

If PatternName is the name of a defined joint pattern, the material temperature for the frame object may vary from one end to the other. The material temperature at each end of the object is equal to the specified temperature multiplied by the pattern value at the joint at the end of the frame object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects and the Name item is ignored.

## Remarks

This function assigns material temperatures to frame objects.

The function returns zero if the material temperatures are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameMatTemp()

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



   'assign material temperature

      ret = SapModel.FrameObj.SetMatTemp("ALL", 50, , Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetMatTemp](GetMatTemp_%7BFrame_Object%7D.md)
