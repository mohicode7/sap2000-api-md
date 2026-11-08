# SetShellLayer (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropArea.SetShellLayer

## VB6 Procedure

Function SetShellLayer(ByVal Name As String, ByVal NumberLayers As Long, ByRef LayerName() As String, ByRef Dist() As Double, ByRef Thickness() As Double, ByRef MatProp() As String, ByRef NonLinear() As Boolean, ByRef MatAng() As Double, ByRef NumIntegrationPts() As Long) As Long

## Parameters

Name

The name of an existing shell-type area property that is specified to be a layered shell property.

NumberLayers

The number of layers in the area property.

LayerName

This is an array that includes the name of each layer.

Dist

This is an array that includes the distance from the area reference surface (area object joint location plus offsets) to the midheight of the layer. [L]

Thickness

This is an array that includes the thickness of each layer. [L]

MatProp

This is an array that includes the name of the material property for the layer.

NonLinear

This is an array that includes a boolean (True or False) value. If this item is True, and if the material property assigned to the layer is nonlinear, the layer will behave nonlinearly in a nonlinear load case. If this item is False, the layer will never behave nonlinearly.

MatAng

This is an array that includes the material angle for the layer. [deg]

NumIntegrationPts

The number of integration points in the thickness direction for the layer. The locations are determined by the program using standard Guass-quadrature rules.

## Remarks

This function assigns the layer parameters for shell-type area properties.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

The function returns an error if the specified area property is not a shell-type property specified to be a layered shell.

## VBA Example



```
Sub SetAreaPropShellLayer()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim MyNumberLayers As Long

      Dim MyLayerName() As String

      Dim MyDist() As Double

      Dim MyThickness() As Double

      Dim MyMatProp() As String

      Dim MyNonLinear() As Boolean

      Dim MyMatAng() As Double

      Dim MyNumIntegrationPts() As Long



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



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

      ReDim MyLayerName(MyNumberLayers - 1)

      ReDim MyDist(MyNumberLayers - 1)

      ReDim MyThickness(MyNumberLayers - 1)

      ReDim MyMatProp(MyNumberLayers - 1)

      ReDim MyNonLinear(MyNumberLayers - 1)

      ReDim MyMatAng(MyNumberLayers - 1)

      ReDim MyNumIntegrationPts(MyNumberLayers - 1)



      MyLayerName(0) = "Concrete"

      MyDist(0) = 0

      MyThickness(0) = 16

      MyMatProp(0) = "4000Psi"

      MyNonLinear(0) = False

      MyMatAng(0) = 0

      MyNumIntegrationPts(0) = 2



      MyLayerName(1) = "Top Bar 1"

      MyDist(1) = 6

      MyThickness(1) = 0.03

      MyMatProp(1) = Name

      MyNonLinear(1) = False

      MyMatAng(1) = 0

      MyNumIntegrationPts(1) = 1



      MyLayerName(2) = "Top Bar 2"

      MyDist(2) = 6

      MyThickness(2) = 0.03

      MyMatProp(2) = Name

      MyNonLinear(2) = False

      MyMatAng(2) = 90

      MyNumIntegrationPts(2) = 1



      MyLayerName(3) = "Bot Bar 1"

      MyDist(3) = -6

      MyThickness(3) = 0.03

      MyMatProp(3) = Name

      MyNonLinear(3) = False

      MyMatAng(3) = 0

      MyNumIntegrationPts(3) = 1



      MyLayerName(4) = "Bot Bar 2"

      MyDist(4) = -6

      MyThickness(4) = 0.03

      MyMatProp(4) = Name

      MyNonLinear(4) = False

      MyMatAng(4) = 90

      MyNumIntegrationPts(4) = 1



      ret = SapModel.PropArea.SetShellLayer("A1", MyNumberLayers, MyLayerName, MyDist, MyThickness, MyMatProp, MyNonLinear, MyMatAng, MyNumIntegrationPts)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

This function is obsolete and has been superseded by [SetShellLayer\_1](../Definitions/Properties/Area/SetShellLayer_1.md) as of version 12.5. This function is maintained for backwards compatibility.

## See Also

[GetShellLayer](GetShellLayer.md)

[GetShellLayer\_1](../Definitions/Properties/Area/GetShellLayer_1.md)
