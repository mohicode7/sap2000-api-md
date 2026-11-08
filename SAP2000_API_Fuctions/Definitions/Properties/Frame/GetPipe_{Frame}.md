# GetPipe

## Syntax

SapObject.SapModel.PropFrame.GetPipe

## VB6 Procedure

Function GetPipe(ByVal Name As String, ByRef FileName As String, ByRef MatProp As String, ByRef t3 As Double, ByRef tw As Double, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing frame section property.

FileName

If the section property was imported from a property file, this is the name of that file. If the section property was not imported, this item is blank.

MatProp

The name of the material property for the section.

t3

The outside diameter. [L]

tw

The wall thickness. [L]

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves frame section property data for a pipe-type frame section.

The function returns zero if the section property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropPipe()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String

      Dim MatProp As String

      Dim t3 As Double

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set new frame section property

      ret = SapModel.PropFrame.SetPipe("PIPE1", "A992Fy50", 8, 0.375)



   'get frame section property data

      ret = SapModel.PropFrame.GetPipe("PIPE1", FileName, MatProp, t3, tw, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetPipe](SetPipe.md)
