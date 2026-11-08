# SetPolygon

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetPolygon

## VB6 Procedure

Function SetPolygon(ByVal Name As String, ByRef ShapeName As String, ByVal MatProp As String, ByVal SSOverwrite As String, ByVal NumberPoints As Long, ByRef X() As Double, ByRef Y() As Double, ByRef Radius() As Double, Optional ByVal Color As Long = -1, Optional ByVal Reinf As Boolean = False, Optional ByVal MatRebar As String = "") As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing or new shape in a section designer property. If this is an existing shape, that shape is modified; otherwise, a new shape is added.

This item may be input as a blank string, in which case the program will assign a shape name to the shape and return that name in the ShapeName variable.

MatProp

The name of the material property for the shape.

SSOverwrite

This is a blank string, Default, or the name of a defined stress-strain curve.

If this item is a blank string or Default, the shape stress-strain curve is based on the assigned material property.

NumberPoints

The number of point coordinates used to define the polygon.

X

This is an array that contains the X-coordinates of the polygon points. [L]

Y

This is an array that contains the Y-coordinates of the polygon points. [L]

Radius

This is an array that contains the radius to be applied at each of the polygon points. [L]

Color

The fill color assigned to the shape. If Color is specified as -1, the program will automatically assign the default fill color.

Reinf

This item is True when there is edge and corner reinforcing steel associated with the shape.

If the MatProp item is not a concrete material, this item is always assumed to be False.

MatRebar

The material property for the edge and corner reinforcing steel associated with the shape. This item applies only when the MatProp item is a concrete material and the Reinf item is True.

## Remarks

This function adds a new polygon shape or modifies an existing shape to be a polygon shape in a section designer property.

The polygon points must be defined in order around the polygon, otherwise the shape will be created incorrectly or the creation of the shape will fail.

The function returns zero if the shape is successfully added or modified; otherwise it returns a nonzero val .

## VBA Example



```
Sub SetFrameSDPropPolygon()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String
      Dim NumberPoints As Long
      Dim X() As Double
      Dim Y() As Double
      Dim Radius() As Double



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



   'add ASTM A706 rebar material

      ret = SapModel.PropMaterial.AddQuick(Name, eMatType.eMatType_Rebar, , , , , eMatTypeRebar.eMatTypeRebar_ASTM_A706)



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", "A992Fy50")



   'add polygon shape to new property

      NumberPoints = 5

      ReDim X(4)

      ReDim Y(4)

      ReDim Radius(4)

      X(0) = -26

      Y(0) = -26

      Radius(0) = 0
      X(1) = -20

      Y(1) = -20

      Radius(1) = 5

      X(2) = -25

      Y(2) = 15

      Radius(2) = 0

      X(3) = 20

      Y(3) = 12

      Radius(3) = 3

      X(4) = 25

      Y(4) = -15

      Radius(4) = 0

      ret = SapModel.PropFrame.SDShape.SetPolygon("SD1", "SH1", "4000Psi", "Default", NumberPoints, X, Y, Radius, -1, True, Name)  
 

   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.1.0.

## See Also

[GetPolygon](GetPolygon.md)
