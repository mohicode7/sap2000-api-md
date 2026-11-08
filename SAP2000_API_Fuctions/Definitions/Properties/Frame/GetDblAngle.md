# GetDblAngle\_2

## Syntax

SapObject.SapModel.PropFrame.GetDblAngle\_2

## VB6 Procedure

Function GetDblAngle\_2(ByVal Name As String, ByRef FileName
As String, ByRef MatProp As String, ByRef t3 As Double, ByRef t2 As Double,
ByRef tf As Double, ByRef tw As Double, ByRef dis As Double, ByRef FilletRadius
As Double, ByRef MirrorAbout3 As Boolean, ByRef Color As Long, ByRef Notes
As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing frame section property.

FileName

If the section property was imported from a property
file, this is the name of that file. If the section property was not imported,
this item is blank.

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

FilletRadius

The fillet radius. [L]

MirrorAbout3 (not applicable)

Indicates whether the section is mirrored about the
local 3-axis.

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned
to the section.

## Remarks

This function retrieves frame section property data
for a double angle-type frame section.

The function returns zero if the section property data
is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropDblAngle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String

      Dim MatProp As String

      Dim t3 As Double

      Dim t2 As Double

      Dim tf As Double

      Dim tw As Double

      Dim dis As Double
```



    Dim
FilletRadius As Double



```
 Dim MirrorAbout3 As 
 Boolean 
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



   'set new frame section property

      ret = SapModel.PropFrame.SetDblAngle_2("DBANG1", 
 "A992Fy50", 6, 9, 0.5, 0.5, 1, 0.2, False)



   'get frame section property data

      ret = SapModel.PropFrame.GetDblAngle_2("DBANG1", 
 FileName, MatProp, t3, t2, tf, tw, dis, FilletRadius, MirrorAbout3, Color, 
 Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 25.0.0. This function replaces
the [GetDblAngle](../../../Obsolete_Functions/GetDblAngle_old.md).

## See Also

[SetDblAngle\_2](SetDblAngle_%7BFrame%7D.md)
