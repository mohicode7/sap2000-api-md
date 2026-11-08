# SetOAluminum

## Syntax

SapObject.SapModel.PropMaterial.SetOAluminum

## VB6 Procedure

Function SetOAluminum(ByVal Name As String, ByVal MyType As Long, ByVal Alloy As String, ByVal Fcy As Double, ByVal Fty As Double, ByVal Ftu As Double, ByVal Fsu As Double, ByVal SSHysType As Long, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing aluminum material property.

MyType

This is 1, 2 or 3, indicating the type of aluminum.

1 = Wrought

2 = Cast-Mold

3 = Cast-Sand

Alloy

The Alloy designation for the aluminum, for example, 2014-T6 for wrought or 356.0-T7 for cast (mold or sand) aluminum.

Fcy

The compressive yield strength of aluminum. [F/L2]

Fty

The tensile yield strength of aluminum. [F/L2]

Ftu

The tensile ultimate strength of aluminum. [F/L2]

Fsu

The shear ultimate strength of aluminum. [F/L2]

SSHysType

This is 0, 1 or 2, indicating the stress-strain hysteresis type.

0 = Elastic

1 = Kinematic

2 = Takeda

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been defined previously for the material.

## Remarks

This function sets the other material property data for aluminum materials.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropAluminumData()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'initialize new material property

      ret = SapModel.PropMaterial.SetMaterial("Aluminum", MATERIAL_ALUMINUM)



   'assign other properties

      ret = SapModel.PropMaterial.SetOAluminum("Aluminum", 1, "2014-T6", 34, 34, 37, 23, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetOAluminum](GetOAluminum.md)
