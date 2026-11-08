# GetAutoSelectSteel

## Syntax

SapObject.SapModel.PropFrame.GetAutoSelectSteel

## VB6 Procedure

Function GetAutoSelectSteel(ByVal Name As String, ByVal NumberItems As Long, ByRef SectName() As String, ByRef AutoStartSection As String, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing auto select section list–type frame section property.

NumberItems

The number of frame section properties included in the auto select list.

SectName

This is an array of the names of the frame section properties included in the auto select list.

AutoStartSection

This is either Median or the name of a frame section property in the SectName array. It is the starting section for the auto select list.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves frame section property data for a steel auto select lists.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropAutoSelectSteel()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyName() As String

      Dim NumberItems As Long

      Dim SectName() As String

      Dim AutoStartSection As String

      Dim Notes As String

      Dim GUID As String



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



   'get auto select list data

      ret = SapModel.PropFrame.GetAutoSelectSteel("AUTO1", NumberItems, SectName, AutoStartSection, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetAutoSelectSteel](SetAutoSelectSteel.md)
