# GetAsolid

## Syntax

SapObject.SapModel.PropArea.GetAsolid

## VB6 Procedure

Function GetAsolid(ByVal Name As String, ByRef MatProp As String, ByRef MatAng As Double, ByRef Arc As Double, ByRef Incompatible As Boolean, ByRef CSys As String, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing asolid-type area property.

MatProp

The name of the material property for the area property.

MatAng

The material angle. [deg]

Arc

The arc angle through which the area object is passed to define the asolid element. [deg]

A value of zero means 1 radian (approximately 57.3 degrees).

Incompatible

If this item is True, incompatible bending modes are included in the stiffness formulation. In general, incompatible modes significantly improve the bending behavior of the object.

CSys

The area object is rotated about the Z-axis in this coordinate system to define the asolid element.

Color

The display color assigned to the property.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property.

## Remarks

This function retrieves area property data for an asolid-type area section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaPropAsolid()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatProp As String

      Dim MatAng As Double

      Dim Arc As Double

      Dim Incompatible As Boolean

      Dim CSys As String

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

      ret = SapModel.PropArea.SetAsolid("AS1", "4000Psi", 0, 45, True)



   'get area property data

      ret = SapModel.PropArea.GetAsolid("AS1", MatProp, MatAng, Arc, Incompatible, CSys, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetAsolid](SetAsolid.md)
