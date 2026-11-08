# GetORebar\_1

## Syntax

SapObject.SapModel.PropMaterial.GetORebar\_1

## VB6 Procedure

Function GetORebar\_1(ByVal Name As String, ByRef Fy As Double, ByRef Fu As Double, ByRef eFy As Double, ByRef eFu As Double, ByRef SSType As Long, ByRef SSHysType As Long, ByRef StrainAtHardening As Double, ByRef StrainUltimate As Double, ByRef FinalSlope As Double, ByRef UseCaltransSSDefaults As Boolean, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing rebar material property.

Fy

The minimum yield stress. [F/L2]

Fu

The minimum tensile stress. [F/L2]

eFy

The expected yield stress. [F/L2]

eFu

The expected tensile stress. [F/L2]

SSType

This is 0, 1 or 2, indicating the stress-strain curve type.

0 = User defined

1 = Parametric - Simple

2 = Parametric - Park

SSHysType

This is 0, 1 or 2, indicating the stress-strain hysteresis type.

0 = Elastic

1 = Kinematic

2 = Takeda

StrainAtHardening

This item applies only when parametric stress-strain curves are used and when UseCaltransSSDefaults is False. It is the strain at the onset of strain hardening.

StrainUltimate

This item applies only when parametric stress-strain curves are used and when UseCaltransSSDefaults is False. It is the ultimate strain capacity. This item must be larger than the StrainAtHardening item.

FinalSlope

This item applies only to parametric stress-strain curves. It is a multiplier on the material modulus of elasticity, E. This value multiplied times E gives the final slope of the curve.

UseCaltransSSDefaults

If this item is True, the program uses Caltrans default controlling strain values, which are bar size dependent.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the other material property data for rebar materials.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value. The function returns an error if the specified material is not rebar.

## VBA Example



```
Sub GetMatPropRebarData_1()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Fy As Double

      Dim Fu As Double

      Dim eFy As Double

      Dim eFu As Double

      Dim SSType As Long

      Dim SSHysType As Long

      Dim StrainAtHardening As Double

      Dim StrainUltimate As Double

      Dim FinalSlope As Double

      Dim UseCaltransSSDefaults As Boolean



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



   'initialize new material property

      ret = SapModel.PropMaterial.SetMaterial("Rebar", MATERIAL_REBAR)



   'assign other properties

      ret = SapModel.PropMaterial.SetORebar_1("Rebar", 62, 93, 70, 102, 2, 2, 0.02, 0.1, -0.1, False)



   'get other properties

      ret = SapModel.PropMaterial.GetORebar_1("Rebar", Fy, Fu, eFy, eFu, SSType, SSHysType, StrainAtHardening, StrainUltimate, FinalSlope, UseCaltransSSDefaults)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

This function supersedes [GetORebar](../../../Obsolete_Functions/GetORebar.md).

## See Also

[SetORebar\_1](SetORebar_1.md)
