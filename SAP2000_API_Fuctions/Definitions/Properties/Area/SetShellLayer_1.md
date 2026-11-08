# SetShellLayer\_1

## Syntax

SapObject.SapModel.PropArea.SetShellLayer\_1

## VB6 Procedure

Function SetShellLayer\_1(ByVal Name As String, ByVal NumberLayers As Long, ByRef LayerName() As String, ByRef Dist() As Double, ByRef Thickness() As Double, ByRef MatProp() As String, ByRef NonLinear() As Boolean, ByRef MatAng() As Double, ByRef NumIntegrationPts() As Long) As Long

## Parameters

Name

The name of an existing shell-type area property that is specified to be a layered shell property.

NumberLayers

The number of layers in the area property.

LayerName

This is an array that includes the name of each layer.

Dist

This is an array that includes the distance from the area reference surface (area object joint location plus offsets) to the mid-height of the layer. [L]

Thickness

This is an array that includes the thickness of each layer. [L]

Type

This is an array that includes 1, 2 or 3, indicating the layer type.

1 = Shell

2 = Membrane

3 = Plate

NumIntegrationPts

The number of integration points in the thickness direction for the layer. The locations are determined by the program using standard Guass-quadrature rules.

MatProp

This is an array that includes the name of the material property for the layer.

MatAng

This is an array that includes the material angle for the layer. [deg]

S11Type, S22Type, S12Type

These are arrays that include 0, 1 or 2, indicating the material component behavior.

0 = Inactive

1 = Linear

2 = Nonlinear

## Remarks

This function assigns the layer parameters for shell-type area properties.

The function returns zero if the parameters are successfully assigned; otherwise, it returns a nonzero value.

The function returns an error if the specified area property is not a shell-type property specified to be a layered shell.

## VBA Example



```
Sub SetAreaPropShellLayer1()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim MyNumberLayers As Long

      Dim MyLayerName() As String

      Dim MyDist() As Double

      Dim MyThickness() As Double

      Dim MyType() As Long

      Dim MyNumIntegrationPts() As Long

      Dim MyMatProp() As String

      Dim MyMatAng() As Double

      Dim MyS11Type() As Long

      Dim MyS22Type() As Long

      Dim MyS12Type() As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart

 
   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'set new area property

      ret = SapModel.PropArea.SetShell("A1", 6, "", 0, 0, 0)



   'add A615Gr60 rebar material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A615Gr60)



   'set area property layer parameters

      MyNumberLayers = 5

      ReDim MyLayerName(MyNumberLayers – 1)

      ReDim MyDist(MyNumberLayers – 1)

      ReDim MyThickness(MyNumberLayers – 1)

      ReDim MyType(MyNumberLayers – 1)

      ReDim MyNumIntegrationPts(MyNumberLayers – 1)

      ReDim MyMatProp(MyNumberLayers – 1)

      ReDim MyMatAng(MyNumberLayers – 1)

      ReDim MyS11Type(MyNumberLayers – 1)

      ReDim MyS22Type(MyNumberLayers – 1)

      ReDim MyS12Type(MyNumberLayers – 1)



      MyLayerName(0) = "Concrete"

      MyDist(0) = 0

      MyThickness(0) = 16

      MyType(0) = 1

      MyNumIntegrationPts(0) = 2

      MyMatProp(0) = "4000Psi"

      MyMatAng(0) = 0

      MyS11Type(0) = 1

      MyS22Type(0) = 1

      MyS12Type(0) = 1



      MyLayerName(1) = "Top Bar 1"

      MyDist(1) = 6

      MyThickness(1) = 0.03

      MyType(1) = 1

      MyNumIntegrationPts(1) = 1

      MyMatProp(1) = Name

      MyMatAng(1) = 0

      MyS11Type(1) = 1

      MyS22Type(1) = 1

      MyS12Type(1) = 1



      MyLayerName(2) = "Top Bar 2"

      MyDist(2) = 6

      MyThickness(2) = 0.03

      MyType(2) = 1

      MyNumIntegrationPts(2) = 1

      MyMatProp(2) = Name

      MyMatAng(2) = 90

      MyS11Type(1) = 1

      MyS22Type(1) = 1

      MyS12Type(1) = 1



      MyLayerName(3) = "Bot Bar 1"

      MyDist(3) = -6

      MyThickness(3) = 0.03

      MyType(3) = 1

      MyNumIntegrationPts(3) = 1

      MyMatProp(3) = Name

      MyMatAng(3) = 0

      MyS11Type(1) = 1

      MyS22Type(1) = 1

      MyS12Type(1) = 1



      MyLayerName(4) = "Bot Bar 2"

      MyDist(4) = -6

      MyThickness(4) = 0.03

      MyType(4) = 1

      MyNumIntegrationPts(4) = 1

      MyMatProp(4) = Name

      MyMatAng(4) = 90

      MyS11Type(1) = 1

      MyS22Type(1) = 1

      MyS12Type(1) = 1



      ret = SapModel.PropArea.SetShellLayer_1("A1", MyNumberLayers, MyLayerName, MyDist, MyThickness, MyType, MyNumIntegrationPts, MyMatProp, MyMatAng, MyS11Type, MyS22Type, MyS12Type)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.5.

This function supersedes [SetShellLayer](../../../Obsolete_Functions/SetShellLayer.md)

## See Also

GetShellLayer\_1
