# SetHybridUSection

## Syntax

SapObject.SapModel.PropFrame.SetHybridUSection

## VB6 Procedure

Function SetHybridUSection(ByVal Name As String, ByRef WebMaterial As String, ByRef TopFlangeMaterial As String, ByRef BotFlangeMaterial As String, ByRef T() As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property then that property is modified, otherwise, a new property is added.

MatPropTopFlange

The name of the material property for the top flange.

MatPropWeb

The name of the material property for the web.

MatPropBotFlange

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

The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, then the program assigns a GUID to the section.

## Remarks

This function initializes a steel hybrid U frame section property. If this function is called for an existing frame section property then all items for the section are reset to their default value.

The function returns zero if the section property is successfully initialized, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropHybridUSection()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim T() As Double

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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also
