# GetHybridUSection

## Syntax

SapObject.SapModel.PropFrame.GetHybridUSection

## VB6 Procedure

Function GetHybridUSection(ByVal name As String, ByRef WebMaterial As String, ByRef TopFlangeMaterial As String, ByRef BotFlangeMaterial As String, ByRef T() As Double, ByRef color As Long, ByRef notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing frame section property.

WebMaterial

The name of the material property for the webs.

TopFlangeMaterial

The name of the material property for the top flanges.

BotFlangeMaterial

The name of the material property for the bottom flange.

T()

The dimension array of the section:

(0) D1 = Web Depth (vertical, inside to inside of flanges). [L]

(1) B1 = Web Distance at Top (CL to CL). [L]

(2) B2 = Bottom Flange Width. [L]

(3) B3 = Top Flange Width (per each). [L]

(4) B4 = Bottom Flange Lip (Web CL to flange edge, may be zero). [L]

(5) tw = Web Thickness. [L]

(6) tf = Top Flange Thickness. [L]

(7) tfb = Bottom Flange Thickness. [L]

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves frame section property data for a steel hybrid U frame section.

The function returns zero if the section property data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropHybridUSection()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim WebMaterial As String

      Dim TopFlangeMaterial As String

      Dim BotFlangeMaterial As String

      Dim T() As Double

      Dim MyT() As Double

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



   'set new Hybrid U-type frame section property

      Redim T(7)

      T(0) = 48

      T(1) = 64

      T(2) = 48

      T(3) = 15

      T(4) = 1.5

      T(5) = 1

      T(6) = 2

      T(7) = 2

      ret = SapModel.PropFrame.SetHybridUSection("HybridU",  "A992Fy50", "A992Fy50", "A992Fy50", T())



   'get frame section property data for cover plated I

      ret = SapModel.PropFrame.GetHybridUSection("HybridI", WebMaterial, TopFlangeMaterial, BotFlangeMaterial, MyT(), Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also
