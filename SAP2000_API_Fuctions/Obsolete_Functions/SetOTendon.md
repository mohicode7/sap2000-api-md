# SetOTendon (Note:  Newer function available)

## Syntax

SapObject.SapModel.PropMaterial.SetOTendon

## VB6 Procedure

Function SetOTendon(ByVal Name As String, ByVal Fy As Double, ByVal Fu As Double, ByVal SSType As Long, ByVal SSHysType As Long, Optional ByVal Temp As Double = 0) As Long

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

This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.

## Remarks

This function sets the other material property data for tendon materials.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropTendonData()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

This function is obsolete and has been superseded by [SetOTendon\_1](../Definitions/Properties/Material/SetOTendon_1.md) as of version 12.00. This function is maintained for backwards compatibility.

## See Also

[GetOTendon](GetOTendon.md)
