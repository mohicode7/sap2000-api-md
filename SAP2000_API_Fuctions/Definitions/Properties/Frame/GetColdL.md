# GetColdL

## Syntax

SapObject.SapModel.PropFrame.GetColdL

## VB6 Procedure

Function GetColdL(ByVal Name As String, ByRef FileName
As String, ByRef MatProp As String, ByRef t3 As Double, ByRef Thickness
As Double, ByRef Radius As Double, ByRef LipDepth As Double, ByRef Color
As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing cold formed Angle frame section
property.

FileName

If the section property was imported from a property
file, this is the name of that file. If the section property was not imported,
this item is blank.

MatProp

The name of the material property for the section.

t3

The section depth. [L]

Thickness

The section thickness. [L]

Radius

The corner radius, if any. [L]

LipDepth

The lip depth, if any. [L]

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned
to the section.

## Remarks

This function retrieves frame section property data
for a cold formed Angle frame section.

The function returns zero if the section property data
is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropColdL()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim FileName As String

      Dim MatProp As String

      Dim t3 As Double

      Dim Thickness As Double

      Dim Radius As Double

      Dim LipDepth As Double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, 
 MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'set new frame section property

      ret = SapModel.PropFrame.SetColdL("CL1", 
 Name, 4.0, 0.08, 0.3, 0.6)



   'get frame section property data

      ret = SapModel.PropFrame.GetColdL("CL1", 
 FileName, MatProp, t3, Thickness , Radius, LipDepth, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 22.2.0.

## See Also

[SetColdL](SetColdL.md)
