# GetRefLine

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetRefLine

## VB6 Procedure

Function GetRefLine(ByVal Name As String, ByVal ShapeName As String, ByRef X1 As Double, ByRef Y1 As Double, ByRef X2 As Double, ByRef Y2 As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of a reference line shape in the section designer section.

X1

The section designer X coordinate of the first drawn end point of the line pattern reinforcing. [L]

Y1

The section designer Y coordinate of the first drawn end point of the line pattern reinforcing. [L]

X2

The section designer X coordinate of the second drawn end point of the line pattern reinforcing. [L]

Y2

The section designer Y coordinate of the second drawn end point of the line pattern reinforcing. [L]

## Remarks

This function retrieves property data for a reference line shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropRefLine()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatRebar As String

      Dim X1 As Double

      Dim Y1 As Double

      Dim X2 As Double

      Dim Y2 As Double



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



   'get Reference Line shape property data

      ret = SapModel.PropFrame.SDShape.GetRefLine("SD1", "SH1", X1, Y1, X2, Y2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetRefLine](SetRefLine.md)
