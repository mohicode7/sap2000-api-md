# GetColdZ

## Syntax

SapObject.SapModel.PropFrame.GetColdZ

## VB6 Procedure

Function GetColdZ(ByVal Name As String, ByRef FileName As String, ByRef MatProp As String, ByRef t3 As Double, ByRef t2 As Double, ByRef Thickness As Double, ByRef Radius As Double, ByRef LipDepth As Double, ByRef LipAngle As Double, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing cold formed Z-type frame section property.

FileName

If the section property was imported from a property file, this is the name of that file. If the section property was not imported, this item is blank.

MatProp

The name of the material property for the section.

t3

The section depth. [L]

t2

The section width. [L]

Thickness

The section thickness. [L]

Radius

The corner radius, if any. [L]

LipDepth

The lip depth, if any. [L]

LipAngle

The lip angle measured from horizontal (0 <= LipAngle <= 90). [deg]

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves frame section property data for a cold formed Z-type frame section.

The function returns zero if the section property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropColdZ()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim FileName As String

      Dim MatProp As String

      Dim t3 As Double

      Dim t2 As Double

      Dim Thickness As Double

      Dim Radius As Double

      Dim LipDepth As Double

      Dim LipAngle As Double

      Dim Color As Long

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



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'set new frame section property

      ret = SapModel.PropFrame.SetColdZ("CZ1", Name, 10, 3.2, 0.08, 0.3, 0.6, 60)



   'get frame section property data

      ret = SapModel.PropFrame.GetColdZ("CZ1", FileName, MatProp, t3, t2, Thickness , Radius, LipDepth, LipAngle, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetColdZ](SetColdZ.md)
