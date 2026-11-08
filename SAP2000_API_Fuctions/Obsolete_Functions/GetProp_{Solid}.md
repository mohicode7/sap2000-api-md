# GetProp (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropSolid.GetProp

## VB6 Procedure

Function GetProp(ByVal Name As String, ByRef MatProp As String, ByRef a As Double, ByRef B As Double, ByRef c As Double, ByRef Incompatible As Boolean, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

Name

The name of an existing solid property.

MatProp

The name of the material property assigned to the solid property.

a

The material angle A. [deg]

b

The material angle B. [deg]

c

The material angle C. [deg]

Incompatible

If this item is True, incompatible bending modes are included in the stiffness formulation. In general, incompatible modes significantly improve the bending behavior of the object.

Color

The display color assigned to the property.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property.

## Remarks

This function retrieves solid property definition data.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSolidProperty()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim MatProp As String

          Dim a As Double

          Dim b As Double

          Dim c As Double

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

          ret = SapModel.File.NewSolidBlock(20, 50, 20)



       'get solid property data

          ret = SapModel.PropSolid.GetProp("SOLID1", MatProp, a, b, c, Incompatible, Color, Notes, GUID)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.02.

The function is obsolete and has been superseded by [GetProp\_1](../Definitions/Properties/Solid/GetProp_1_%7BSolid%7D.md) as of version 26.40. This function is maintained for backward compatibility, but the parameter Incompatible is no longer applicable and will always be returned as false.

## See Also

[SetProp\_1](../Definitions/Properties/Solid/SetProp_1_%7BSolid%7D.md)

[GetProp\_1](../Definitions/Properties/Solid/GetProp_1_%7BSolid%7D.md)
