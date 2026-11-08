# SetSection

## Syntax

SapObject.SapModel.FrameObj.SetSection

## VB6 Procedure

Function SetSection(ByVal name As String, ByVal PropName As String, Optional ByVal ItemType As eItemType = object, Optional ByVal sVarRelStartLoc As Double = 0, Optional ByVal sVarTotalLength As Double = 0) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

PropName

This is None or the name of a frame section property to be assigned to the specified frame object(s).

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

sVarTotalLength

This is the total assumed length of the nonprismatic section. Enter 0 for this item to indicate that the section length is the same as the frame object length.

This item is applicable only when the assigned frame section property is a nonprismatic section.

sVarRelStartLoc

This is the relative distance along the nonprismatic section to the I-End (start) of the frame object. This item is ignored when the sVarTotalLengthitem is 0.

This item is applicable only when the assigned frame section property is a nonprismatic section, and the sVarTotalLengthitem is greater than zero.

## Remarks

This function assigns a frame section property to a frame object.

The function returns zero if the frame section property data is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSectionProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open an existing file

      FileName = "C:\SapAPI\Example 1-022.sdb"

      ret = SapModel.File.OpenFile(FileName)



   'unlock model

      ret = SapModel.SetModelIsLocked(False)



   'set frame section property

      ret = SapModel.FrameObj.SetSection("28", "W24X160")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSection](GetSection_%7BFrame_Object%7D.md)

[GetSectionNonPrismatic](GetSectionNonPrismatic.md)
