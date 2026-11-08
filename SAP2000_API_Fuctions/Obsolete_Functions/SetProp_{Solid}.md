# SetProp (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropSolid.SetProp

## VB6 Procedure

Function SetProp(ByVal Name As String, ByVal MatProp As String, ByVal a As Double, ByVal B As Double, ByVal c As Double, ByVal Incompatible As Boolean, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

## Parameters

Name

The name of an existing or new solid property. If this is an existing property, that property is modified; otherwise, a new property is added.

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

The display color assigned to the property. If Color is specified as -1, the program will automatically assign a color.

Notes

The notes, if any, assigned to the property.

GUID

The GUID (global unique identifier), if any, assigned to the property. If this item is input as Default, the program assigns a GUID to the property.

## Remarks

This function defines a solid property.

The function returns zero if the property is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSolidProperty()

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

          ret = SapModel.File.NewSolidBlock(20, 50, 20)



       'set new solid property

          ret = SapModel.PropSolid.SetProp("S1", "4000Psi", 0, 10, 20, True)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.02.

The function is obsolete and has been superseded by [SetProp\_1](../Definitions/Properties/Solid/SetProp_1_%7BSolid%7D.md) as of version 26.40. This function is maintained for backward compatibility, but the parameter Incompatible is no longer applicable so its value will not be applied to the property.

## See Also

[SetProp\_1](../Definitions/Properties/Solid/SetProp_1_%7BSolid%7D.md)

[GetProp\_1](../Definitions/Properties/Solid/GetProp_1_%7BSolid%7D.md)
