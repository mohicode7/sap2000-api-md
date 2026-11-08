# GetOColdFormed

## Syntax

SapObject.SapModel.PropMaterial.GetOColdFormed

## VB6 Procedure

Function GetOColdFormed(ByVal Name As String, ByRef Fy As Double, ByRef Fu As Double, ByRef SSHysType As Long, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing cold formed material property.

Fy

The minimum yield stress. [F/L2]

Fu

The minimum tensile stress. [F/L2]

SSHysType

This is 0, 1 or 2, indicating the stress-strain hysteresis type.

0 = Elastic

1 = Kinematic

2 = Takeda

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the other material property data for cold formed materials.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value. The function returns an error if the specified material is not cold formed.

## VBA Example



```
Sub GetMatPropColdFormedData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Fy As Double

      Dim Fu As Double

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

      ret = SapModel.PropMaterial.SetMaterial("ColdFormed", MATERIAL_COLDFORMED)



   'assign other properties

      ret = SapModel.PropMaterial.SetOColdFormed("ColdFormed", 52, 67, 1)



   'get other properties

      ret = SapModel.PropMaterial.GetOColdFormed("ColdFormed", Fy, Fu, SSHysType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetOColdFormed](SetOColdFormed.md)
