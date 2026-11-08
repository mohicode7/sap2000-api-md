# GetOAluminum

## Syntax

SapObject.SapModel.PropMaterial.GetOAluminum

## VB6 Procedure

Function GetOAluminum(ByVal Name As String, ByVal MyType As Long, ByRef Alloy As String, ByRef Fcy As Double, ByRef Fty As Double, ByRef Ftu As Double, ByRef Fsu As Double, ByRef SSHysType As Long, Optional ByVal Temp As Double = 0) As Long

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

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the other material property data for aluminum materials.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value. The function returns an error if the specified material is not aluminum.

## VBA Example



```
Sub GetMatPropAluminumData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long

      Dim Alloy As String

      Dim Fcy As Double

      Dim Fty As Double

      Dim Ftu As Double

      Dim Fsu As Double

      Dim SSHysType As Long



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



   'get other properties

      ret = SapModel.PropMaterial.GetOAluminum("Aluminum", MyType, Alloy, Fcy, Fty, Ftu, Fsu, SSHysType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetOAluminum](SetOAluminum.md)
