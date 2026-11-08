# SetAutoSelectAluminum

## Syntax

SapObject.SapModel.PropFrame.SetAutoSelectAluminum

## VB6 Procedure

Function SetAutoSelectAluminum(ByVal Name As String, ByVal NumberItems As Long, ByRef SectName() As String, Optional ByVal AutoStartSection As String = "Median", Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.

NumberItems

The number of frame section properties included in the auto select list.

SectName

This is an array of the names of the frame section properties included in the auto select list.

Auto select lists and nonprismatic (variable) sections are not allowed in this array.

AutoStartSection

This is Median or the name of a frame section property in the SectName array. It is the starting section for the auto select list.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.

## Remarks

This function assigns frame section properties to an auto select list. If this function is called for an existing frame section property, all items for the section are reset to their default value.

The function returns zero if the auto select list is successfully filled; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropAutoSelectAluminum()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim MyName() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'add aluminum material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_ALUMINUM, , , MATERIAL_ALUMINUM_SUBTYPE_6061_T6)



   'create new aluminum frame section properties

      ret = SapModel.PropFrame.SetISection("AI", Name , 18, 6, 0.5, 0.3, 6, 0.5)

      ret = SapModel.PropFrame.SetISection("AI2", Name , 18, 6, 0.6, 0.3, 6, 0.6)

      ret = SapModel.PropFrame.SetISection("AI3", Name , 18, 6, 0.7, 0.3, 6, 0.7)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "AI", "AI")



   'define new auto select list frame section property

      ReDim MyName(2)

      MyName(0) = "AI"

      MyName(1) = "AI2"

      MyName(2) = "AI3"

      ret = SapModel.PropFrame.SetAutoSelectAluminum("AUTO1", 3, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetAutoSelectAluminum](GetAutoSelectAluminum.md)
