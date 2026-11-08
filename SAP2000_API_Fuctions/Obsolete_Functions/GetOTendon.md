# GetOTendon (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropMaterial.GetOTendon

## VB6 Procedure

Function GetOTendon(ByVal Name As String, ByRef Fy As Double, ByRef Fu As Double, ByRef SSType As Long, ByRef SSHysType As Long, Optional ByVal Temp As Double = 0) As Long

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

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the other material property data for tendon materials.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value. The function returns an error if the specified material is not tendon.

## VBA Example



```
Sub GetMatPropTendonData()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Fy As Double

      Dim Fu As Double

      Dim SSType As Long

      Dim SSHysType As Long



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



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

      ret = SapModel.PropMaterial.SetOTendon("Tendon", 230, 255, 1, 1)



   'get other properties

      ret = SapModel.PropMaterial.GetOTendon("Tendon", Fy, Fu, SSType, SSHysType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

This function is obsolete and has been superseded by [GetOTendon\_1](../Definitions/Properties/Material/GetOTendon_1.md) as of version 12.00. This function is maintained for backwards compatibility.

## See Also

[SetOTendon](SetOTendon.md)
