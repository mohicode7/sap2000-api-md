# GetOTendon\_1

## Syntax

SapObject.SapModel.PropMaterial.GetOTendon\_1

## VB6 Procedure

Function GetOTendon\_1(ByVal Name As String, ByRef Fy As Double, ByRef Fu As Double, ByRef SSType As Long, ByRef SSHysType As Long, ByRef FinalSlope As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing tendon material property.

Fy

The minimum yield stress. [F/L2]

Fu

The minimum tensile stress. [F/L2]

SSType

This is 0, 1 or 2, indicating the stress-strain curve type.

0 = User defined

1 = Parametric – 250 ksi strand

2 = Parametric – 270 ksi strand

SSHysType

This is 0, 1 or 2, indicating the stress-strain hysteresis type.

0 = Elastic

1 = Kinematic

2 = Takeda

FinalSlope

This item applies only to parametric stress-strain curves. It is a multiplier on the material modulus of elasticity, E. This value multiplied times E gives the final slope of the curve.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the other material property data for tendon materials.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value. The function returns an error if the specified material is not tendon.

## VBA Example



```
Sub GetMatPropTendonData_1()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Fy As Double

      Dim Fu As Double

      Dim SSType As Long

      Dim SSHysType As Long

      Dim FinalSlope As Double



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

      ret = SapModel.PropMaterial.SetMaterial("Tendon", MATERIAL_TENDON)



   'assign other properties

      ret = SapModel.PropMaterial.SetOTendon_1("Tendon", 230, 255, 1, 1, -0.1)



   'get other properties

      ret = SapModel.PropMaterial.GetOTendon_1("Tendon", Fy, Fu, SSType, SSHysType, FinalSlope)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

This function supersedes [GetOTendon](../../../Obsolete_Functions/GetOTendon.md).

## See Also

[SetOTendon\_1](SetOTendon_1.md)
