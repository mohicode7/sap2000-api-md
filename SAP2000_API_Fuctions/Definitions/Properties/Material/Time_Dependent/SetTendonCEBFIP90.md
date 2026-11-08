# SetTendonCEBFIP90

## Syntax

SapObject.SapModel.PropMaterial.TimeDep.SetTendonCEBFIP90

## VB6 Procedure

Function SetTendonCEBFIP90(ByVal Name As String, ByVal ConsiderSteelRelaxation As Boolean, ByVal CEBFIPClass As Long, ByVal UseSeries As Long, ByVal NumberSeriesTerms As Long, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing tendon material property.

ConsiderSteelRelaxation

If this item is True, time dependence is considered for tendon steel relaxation.

CEBFIPClass

This is either 1 or 2, indicating the CEB FIP-90 class. This item applies only when ConsiderSteelRelaxation = True.

UseSeries

This is either 0 or 1, indicating the steel relaxation integration type.

0 = Full integration

1 = Dirichlet series

This item applies only when ConsiderSteelRelaxation = True.

NumberSeriesTerms

This is the number of series terms used when integrating based on a Dirichlet series. This item applies only when ConsiderSteelRelaxation = True and UseSeries = 1.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been previously defined for the material.

## Remarks

This function sets the time dependent CEB FIP-90 material property data for tendon materials.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropTendonTimeDepCEBFIP90()

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

      ret = SapModel.PropMaterial.SetMaterial("Tendon", MATERIAL_TENDON)



   'assign CEB FIP-90 time dependent data

      ret = SapModel.PropMaterial.TimeDep.SetTendonCEBFIP90("Tendon", True, 2, 1, 12)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetTendonCEBFIP90](GetTendonCEBFIP90.md)
