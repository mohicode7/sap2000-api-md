# GetONoDesign

## Syntax

SapObject.SapModel.PropMaterial.GetONoDesign

## VB6 Procedure

Function GetONoDesign(ByVal Name As String, ByRef FrictionAngle As Double, ByRef DilatationalAngle As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing concrete material property.

FrictionAngle

The Drucker-Prager friction angle, 0 <= FrictionAngle < 90. [deg]

DilatationalAngle

The Drucker-Prager dilatational angle, 0 <= DilatationalAngle < 90. [deg]

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function retrieves the other material property data for no design type materials.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value. The function returns an error if the specified material is not concrete.

## VBA Example



```
Sub GetMatPropNoDesignData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

   Dim ret As Long

      Dim FrictionAngle As Double

      Dim DilatationalAngle As Double



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

      ret = SapModel.PropMaterial.SetMaterial("NoDesign", MATERIAL_NODESIGN)



   'assign other properties

      ret = SapModel.PropMaterial.SetONoDesign("NoDesign", 10, 15)



   'get other properties

      ret = SapModel.PropMaterial.GetONoDesign("NoDesign", FrictionAngle, DilatationalAngle)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetONoDesign](SetONoDesign.md)
