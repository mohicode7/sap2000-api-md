# GetSDSection

## Syntax

SapObject.SapModel.PropFrame.GetSDSection

## VB6 Procedure

Function GetSDSection(ByVal Name As String, ByRef MatProp As String, ByRef NumberItems As Long, ByRef ShapeName() As String, ByRef MyType() As Long, ByRef DesignType As Long, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing section designer property.

MatProp

The name of the base material property for the section.

NumberItems

The number of shapes defined in the section designer section.

ShapeName

This is an array that includes the name of each shape in the section designer section.

MyType

This is an array that includes the type of each shape in the section designer section.

1 = I-section

2 = Channel

3 = Tee

4 = Angle

5 = Double Angle

6 = Box

7 = Pipe

8 = Plate

101 = Solid Rectangle

102 = Solid Circle

103 = Solid Segment

104 = Solid Sector

201 = Polygon

301 = Reinforcing Single

302 = Reinforcing Line

303 = Reinforcing Rectangle

304 = Reinforcing Circle

401 = Reference Line

402 = Reference Circle

501 = Caltrans Square

502 = Caltrans Circle

503 = Caltrans Hexagon

504 = Caltrans Octagon

DesignType

This is 0, 1, 2 or 3, indicating the design option for the section.

0 = No design

1 = Design as general steel section

2 = Design as a concrete column; check the reinforcing

3 = Design as a concrete column; design the reinforcing

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves section property data for a section designer section.

The function returns zero if the section property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropSDSection()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatProp As String

      Dim NumberItems As Long

      Dim ShapeName() As String

      Dim MyType() As Long

      Dim DesignType As Long

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", "A992Fy50")



   'add I-section shapes to new property

      ret = SapModel.PropFrame.SDShape.SetISection("SD1", "SH1", "A992Fy50", "", 0, -9, 0, -1, 18, 6, 1, 0.5, 6, 1)

      ret = SapModel.PropFrame.SDShape.SetISection("SD1", "SH2", "A992Fy50", "", -9, -9, 0, -1, 18, 6, 1, 0.5, 6, 1)

      ret = SapModel.PropFrame.SDShape.SetISection("SD1", "SH3", "A992Fy50", "", 9, -9, 0, -1, 18, 6, 1, 0.5, 6, 1)



   'get section designer section property data

      ret = SapModel.PropFrame.GetSDSection("SD1", MatProp, NumberItems, ShapeName, MyType, DesignType, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetSDSection](SetSDSection.md)
