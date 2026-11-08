# SetProp\_1

## Syntax

SapObject.SapModel.PropSolid.SetProp\_1

## VB6 Procedure

Function SetProp\_1(ByVal Name As String, ByVal MatProp As String, ByVal a As Double, ByVal B As Double, ByVal c As Double, ByVal FormType As eSolidPropFormulationType, ByVal IsNonlinear As Boolean, Optional ByVal Color As Long = -1, Optional ByVal Notes As String = "", Optional ByVal GUID As String = "") As Long

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

FormType

The type of stiffness formulation used for the solid property. The value is one of the following items in the eSolidPropFormulationType enumeration.

               LegacyStandard = 0

               LegacyQuadratic = 1

               Standard = 2

               Quadratic = 3

IsNonlinear

If this item is True, nonlinear material behavior are included in the Solid behavior for nonlinear static and direct-integration time-history analyses. This option is not available for Quadratic and Legacy Quadratic formulation types.

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

          ret = SapModel.PropSolid.SetProp_1("S1", "4000Psi", 0, 10, 20, eSolidPropFormulationType.Standard, False)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 26.40.

This function supersedes [SetProp](../../../Obsolete_Functions/SetProp_%7BSolid%7D.md).

## See Also

[GetProp\_1](GetProp_1_%7BSolid%7D.md)
