# SetTendonScaleFactors

## Syntax

SapObject.SapModel.PropMaterial.TimeDep.SetTendonScaleFactors

## VB6 Procedure

Function SetTendonScaleFactors(ByVal Name As String, ByRef ScaleFactorRelaxtion As Double, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing tendon material property.

ScaleFactorRelaxation

This value multiplies the relaxation coefficient, and hence the relaxation strain, computed for the material during a time-dependent analysis. It has no effect for load cases that do not consider time-dependent effects, or for materials that do not consider relaxation effects. The default value is unity, and the specified value must be positive.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been previously defined for the material.

## Remarks

This function sets scale factors for the time-dependent material property data for tendon materials. If this function is not called, default values of unity are assumed for all scale factors.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropTendonTimeDepCEBIFP90

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

      ret = SapModel.PropMaterial.SetMaterial("Tendon", eMatType,Tendon)



   'assign CEB FIP-90 time dependent data

      ret = SapModel.PropMaterial.TimeDep.SetTendonCEBFIP90("Tendon", True, 2, 1, 12)
 
   'assign time dependent scale factors

      ret = SapModel.PropMaterial.TimeDep.SetTendonScaleFactors("Tendon", 1.15)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.2.0.

## See Also

[GetTendonScaleFactors](GetTendonScaleFactors.md)
