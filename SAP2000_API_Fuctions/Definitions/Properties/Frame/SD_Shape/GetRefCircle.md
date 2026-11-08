# GetRefCircle

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetRefCircle

## VB6 Procedure

Function GetRefCircle(ByVal Name As String, ByVal ShapeName As String, ByRef XCenter As Double, ByRef YCenter As Double, ByRef Diameter As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing or new shape in a section designer property. If this is an existing shape, that shape is modified; otherwise, a new shape is added.

This item may be input as a blank string, in which case the program will assign a shape name to the shape and return that name in the ShapeName variable.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

Diameter

The diameter of the circular shape. [L]

## Remarks

This function retrieves property data for a reference circle shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropRefCircle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim XCenter As Double

      Dim YCenter As Double

      Dim Diameter As Double

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

      ret = SapModel.PropFrame.SetSDSection("SD1", "4000Psi")



   'add reference circle shape to new property

      ret = SapModel.PropFrame.SDShape.SetRefCircle("SD1", "SH1", 5, 5, 120)



   'get reference circle shape property data

      ret = SapModel.PropFrame.SDShape.GetRefCircle("SD1", "SH1", XCenter, YCenter, Diameter)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetRefCircle](SetRefCircle.md)
