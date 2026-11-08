# SetDblAngle (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropFrame.SetDblAngle

## VB6 Procedure

Function SetDblAngle(ByVal Name As String, ByVal MatProp
As String, ByVal t3 As Double, ByVal t2 As Double, ByVal tf As Double,
ByVal tw As Double, ByVal dis As Double, Optional ByVal Color As Long
= -1, Optional ByVal Notes As String = "", Optional ByVal GUID
As String = "") As Long

## Parameters

Name

The name of an existing or new frame section property.
If this is an existing property, that property is modified; otherwise,
a new property is added.

MatProp

The name of the material property for the section.

t3

The vertical leg depth. [L]

t2

The total width of the section, that is, the sum of
the widths of each horizontal leg plus the back-to-back distance. [L]

tf

The horizontal leg thickness. [L]

tw

The vertical leg thickness. [L]

dis

The back-to-back distance between the angles. [L]

Color

The display color assigned to the section. If Color
is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned
to the section. If this item is input as Default, the program assigns
a GUID to the section.

## Remarks

This function initializes a double angle-type frame
section property. If this function is called for an existing frame section
property, all items for the section are reset to their default value.

The function returns zero if the section property is
successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFramePropDblAngle()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'set new frame section property

      ret = SapModel.PropFrame.SetDblAngle("DBANG1", 
 "A992Fy50", 6, 9, 0.5, 0.5, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetDblAngle](../Definitions/Properties/Frame/GetDblAngle.md)
