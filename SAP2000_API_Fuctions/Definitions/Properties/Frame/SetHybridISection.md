# SetHybridISection

## Syntax

SapObject.SapModel.PropFrame.SetHybridISection

## VB6 Procedure

Function SetHybridISection(ByVal Name As String, ByRef MatPropTopFlange As String, ByRef MatPropWeb As String, ByRef MatPropBotFlange As String, ByRef t3 As Double, ByRef t2 As Double, ByRef TF As Double, ByRef TW As Double, ByRef t2b As Double, ByRef tfb As Double, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property. If this is an existing property then that property is modified, otherwise, a new property is added.

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

The display color assigned to the section. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, then the program assigns a GUID to the section.

## Remarks

This function initializes a steel hybrid I frame section property. If this function is called for an existing frame section property then all items for the section are reset to their default value.

The function returns zero if the section property is successfully initialized, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropHybridISection()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

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



   'set new Hybrid I-type frame section property

      ret = SapModel.PropFrame.SetHybridISection("HybridI", "A992Fy50", "A992Fy50", "A992Fy50", 24, 8, 0.5, 0.3, 8, 0.5)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also
