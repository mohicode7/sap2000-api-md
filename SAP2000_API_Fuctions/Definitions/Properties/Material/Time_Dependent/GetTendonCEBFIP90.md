# GetTendonCEBFIP90

## Syntax

SapObject.SapModel.PropMaterial.GetTendonCEBFIP90

## VB6 Procedure

Function GetTendonCEBFIP90(ByVal Name As String, ByRef ConsiderSteelRelaxation As Boolean, ByRef CEBFIPClass As Long, ByRef UseSeries As Long, ByRef NumberSeriesTerms As Long, Optional ByVal Temp As Double = 0) As Long

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

This item is the temperature at which the specified data is to be retrieved. The temperature must have been previously defined for the material.

## Remarks

This function retrieves the time dependent CEB FIP-90 material property data for tendon materials.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value. The function returns an error if the specified material is not tendon.

## VBA Example



```
Sub GetMatPropTendonTimeDepCEBFIP90()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ConsiderSteelRelaxation As Boolean

      Dim CEBFIPClass As Long

      Dim UseSeries As Long

      Dim NumberSeriesTerms As Long



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



   'get CEB FIP-90 time dependent data

      ret = SapModel.PropMaterial.TimeDep.GetTendonCEBFIP90("Tendon", ConsiderSteelRelaxation, CEBFIPClass, UseSeries, NumberSeriesTerms)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetTendonCEBFIP90](SetTendonCEBFIP90.md)
