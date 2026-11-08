# SetRefLine

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetRefLine

## VB6 Procedure

Function SetRefLine(ByVal Name As String, ByRef ShapeName As String, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing or new shape in a section designer property. If this is an existing shape, that shape is modified; otherwise, a new shape is added.

This item may be input as a blank string, in which case the program will assign a shape name to the shape and return that name in the ShapeName variable.

X1

The section designer X coordinate of the first drawn end point of the line pattern reinforcing. [L]

Y1

The section designer Y coordinate of the first drawn end point of the line pattern reinforcing. [L]

X2

The section designer X coordinate of the second drawn end point of the line pattern reinforcing. [L]

Y2

The section designer Y coordinate of the second drawn end point of the line pattern reinforcing. [L]

## Remarks

This function adds a new reference line shape or modifies an existing shape to be a reference line shape in a section designer property.

The function returns zero if the shape is successfully added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropRefLine()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



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



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", "4000psi")



   'add Line Pattern Reinforcing shape to new property

      ret = SapModel.PropFrame.SDShape.SetRefLine("SD1", "SH1", 5, 5, 2, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetRefLine](GetRefLine.md)
