# GetHybridISection

## Syntax

SapObject.SapModel.PropFrame.GetHybridISection

## VB6 Procedure

Function GetHybridISection(ByVal name As String, ByRef MatPropTopFlange As String, ByRef MatPropWeb As String, ByRef MatPropBotFlange As String, ByRef t3 As Double, ByRef t2 As Double, ByRef TF As Double, ByRef TW As Double, ByRef t2b As Double, ByRef tfb As Double,

ByRef color As Long, ByRef notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing frame section property.

MatPropTopFlange

The name of the material property for the top flange.

MatPropWeb

The name of the material property for the web.

MatPropBotFlange

The name of the material property for the bottom flange.

t3

The height of the section. [L]

t2

The width of the top flange. [L]

TF

The thickness of the top flange. [L]

TW

The thickness of the web. [L]

t2b

The width of the bottom flange. [L]

tfb

The thickness of the bottom flange. [L]

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves frame section property data for a steel hybrid I frame section.

The function returns zero if the section property data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropHybridISection()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatPropTopFlange As String

      Dim MatPropWeb As String

      Dim MatPropBotFlange As String

      Dim t3 As Double

      Dim t2 As Double

      Dim TF As Double

      Dim TW As Double

      Dim t2b As Double

      Dim tfb As Double

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



   'set new Hybrid I-type frame section property

      ret = SapModel.PropFrame.SetHybridISection("HybridI", "A992Fy50", "A992Fy50", "A992Fy50", 24, 8, 0.5, 0.3, 8, 0.5)



   'get frame section property data for cover plated I

      ret = SapModel.PropFrame.GetHybridISection("HybridI", MatPropTopFlange, MatPropWeb, MatPropBotFlange, t3, t2, TF,  TW, t2b, tfb, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also
