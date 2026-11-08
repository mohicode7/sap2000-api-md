# SetAutoSelectSteel

## Syntax

SapObject.SapModel.PropFrame.SetAutoSelectSteel

## VB6 Procedure

Function SetAutoSelectSteel(ByVal Name As String, ByVal NumberItems As Long, ByRef SectName() As String, Optional ByVal AutoStartSection As String = "Median", Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.

NumberItems

The number of frame section properties included in the auto select list.

SectName

This is an array of the names of the frame section properties included in the auto select list.

Auto select lists and nonprismatic (variable) sections are not allowed in this array.

AutoStartSection

This is either Median or the name of a frame section property in the SectName array. It is the starting section for the auto select list.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.

## Remarks

This function assigns frame section properties to an auto select list. If this function is called for an existing frame section property, all items for the section are reset to their default value.

The function returns zero if the auto select list is successfully filled; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropAutoSelectSteel()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyName() As String



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



   'set new I-type frame section property

      ret = SapModel.PropFrame.SetISection("ISEC1", "A992Fy50", 24, 8, 0.5, 0.3, 8, 0.5)



   'set new I-type frame section property

      ret = SapModel.PropFrame.SetISection("ISEC2", "A992Fy50", 20, 8, 0.5, 0.3, 8, 0.5)



   'set new I-type frame section property

      ret = SapModel.PropFrame.SetISection("ISEC3", "A992Fy50", 16, 8, 0.5, 0.3, 8, 0.5)



   'set new auto select list frame section property

      ReDim MyName(2)

      MyName(0) = "ISEC1"

      MyName(1) = "ISEC2"

      MyName(2) = "ISEC3"

      ret = SapModel.PropFrame.SetAutoSelectSteel("AUTO1", 3, MyName, "ISEC2")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetAutoSelectSteel](GetAutoSelectSteel.md)
