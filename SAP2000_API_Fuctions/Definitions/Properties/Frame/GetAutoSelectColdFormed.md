# GetAutoSelectColdFormed

## Syntax

SapObject.SapModel.PropFrame.GetAutoSelectColdFormed

## VB6 Procedure

Function GetAutoSelectColdFormed(ByVal Name As String, ByVal NumberItems As Long, ByRef SectName() As String, ByRef AutoStartSection As String, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing auto select section list-type frame section property.

NumberItems

The number of frame section properties included in the auto select list.

SectName

This is an array of the names of the frame section properties included in the auto select list.

AutoStartSection

This is Median or the name of a frame section property in the SectName array. It is the starting section for the auto select list.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves frame section property data for a cold formed auto select lists.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropAutoSelectColdFormed()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

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



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_COLDFORMED, , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section properties

      ret = SapModel.PropFrame.SetISection("CI", Name , 18, 6, 0.5, 0.3, 6, 0.5)

      ret = SapModel.PropFrame.SetISection("CI2", Name , 18, 6, 0.6, 0.3, 6, 0.6)

      ret = SapModel.PropFrame.SetISection("CI3", Name , 18, 6, 0.7, 0.3, 6, 0.7)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "CI", "CI")



   'define new auto select list frame section property

      ReDim MyName(2)

      MyName(0) = "CI"

      MyName(1) = "CI2"

      MyName(2) = "CI3"

      ret = SapModel.PropFrame.SetAutoSelectColdFormed("AUTO1", 3, MyName)



   'get auto select list data

      ret = SapModel.PropFrame.GetAutoSelectColdFormed("AUTO1", NumberItems, SectName, AutoStartSection, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetAutoSelectColdFormed](SetAutoSelectColdFormed.md)
