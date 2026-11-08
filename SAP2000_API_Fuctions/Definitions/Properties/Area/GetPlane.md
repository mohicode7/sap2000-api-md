# GetPlane

## Syntax

SapObject.SapModel.PropArea.GetPlane

## VB6 Procedure

Function GetPlane(ByVal Name As String, ByRef MyType As Long, ByRef MatProp As String, ByRef MatAng As Double, ByRef Thickness As Double, ByRef Incompatible As Boolean, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing plane-type area property.

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

The display color assigned to the property.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property.

## Remarks

This function retrieves area property data for a plane-type area section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaPropPlane()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long

      Dim MatProp As String

      Dim MatAng As Double

      Dim Thickness As Double

      Dim Incompatible As Boolean

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'set new area property

      ret = SapModel.PropArea.SetPlane("P1", 1, "4000Psi", 0, 9, True)



   'get area property data

      ret = SapModel.PropArea.GetPlane("P1", MyType, MatProp, MatAng, Thickness, Incompatible, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetPlane](SetPlane.md)
