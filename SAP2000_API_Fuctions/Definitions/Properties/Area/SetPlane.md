# SetPlane

## Syntax

SapObject.SapModel.PropArea.SetPlane

## VB6 Procedure

Function SetPlane(ByVal Name As String, ByVal MyType As Long, ByVal MatProp As String, ByVal MatAng As Double, ByVal Thickness As Double, ByVal Incompatible As Boolean, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new area property. If this is an existing property, that property is modified; otherwise, a new property is added.

MyType

This is either 1 or 2, indicating the plane type.

1 = Plane-stress

2 = Plane-strain

MatProp

The name of the material property for the area property.

MatAng

The material angle. [deg]

Thickness

The plane thickness. [L]

Incompatible

If this item is True, incompatible bending modes are included in the stiffness formulation. In general, incompatible modes significantly improve the bending behavior of the object.

Color

The display color assigned to the property. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property. If this item is input as Default, the program assigns a GUID to the property.

## Remarks

This function initializes a plane-type area property. If this function is called for an existing area property, all items for the property are reset to their default value.

The function returns zero if the property is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAreaPropPlane()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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

      ret = SapModel.PropArea.SetPlane("P1", 1, "4000Psi", 0, 9, True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetPlane](GetPlane.md)
