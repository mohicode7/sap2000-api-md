# GetProp\_1

## Syntax

SapObject.SapModel.PropSolid.GetProp\_1

## VB6 Procedure

Function GetProp\_1(ByVal Name As String, ByRef MatProp As String, ByRef a As Double, ByRef B As Double, ByRef c As Double, ByRef FormType As eSolidPropFormulationType, ByRef IsNonlinear As Boolean, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

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

FormType

The type of stiffness formulation used for the solid property. The value is one of the following items in the eSolidPropFormulationType enumeration.

               LegacyStandard = 0

               LegacyQuadratic = 1

               Standard = 2

               Quadratic = 3

IsNonlinear

If this item is True, nonlinear material behavior are included in the Solid behavior for nonlinear static and direct-integration time-history analyses. This option is not available for Quadratic and Legacy Quadratic formulation types.

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
      Dim FormType As eSolidPropFormulationType
      Dim IsNonlinear As Boolean
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

          ret = SapModel.PropSolid.GetProp_1("SOLID1", MatProp, a, b, c, FormType, IsNonlinear, Color, Notes, GUID)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 26.40.

This function supersedes [GetProp](../../../Obsolete_Functions/GetProp_%7BSolid%7D.md).

## See Also

[SetProp\_1](SetProp_1_%7BSolid%7D.md)
