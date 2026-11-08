# GetTee (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropFrame.GetTee

## VB6 Procedure

Function GetTee(ByVal Name As String, ByRef FileName
As String, ByRef MatProp As String, ByRef t3 As Double, ByRef t2 As Double,
ByRef tf As Double, ByRef tw As Double, ByRef Color As Long, ByRef Notes
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

The section depth. [L]

t2

The flange width. [L]

tf

The flange thickness. [L]

tw

The web thickness. [L]

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned
to the section.

## Remarks

This function retrieves frame section property data
for a tee-type frame section.

The function returns zero if the section property data
is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropTee()

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

      ret = SapModel.PropFrame.SetTee("TEE1", 
 "A992Fy50", 12, 10, 0.6, 0.3)



   'get frame section property data

      ret = SapModel.PropFrame.GetTee("TEE1", 
 FileName, MatProp, t3, t2, tf, tw, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetTee](../Definitions/Properties/Frame/GetTee_%7BFrame%7D.md)

[SetRebarBeam](../Definitions/Properties/Frame/SetRebarBeam.md)

[GetRebarBeam](../Definitions/Properties/Frame/GetRebarBeam.md)
