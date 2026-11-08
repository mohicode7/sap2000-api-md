# SetConcreteCEBFIP90

## Syntax

SapObject.SapModel.PropMaterial.TimeDep.SetConcreteCEBFIP90

## VB6 Procedure

Function SetConcreteCEBFIP90(ByVal Name As String, ByVal ConsiderConcreteAge As Boolean, ByVal ConsiderConcreteCreep As Boolean, ByVal ConsiderConcreteShrinkage As Boolean, ByVal CEBFIPsCoefficient As Double, ByVal RelativeHumidity As Double, ByVal NotionalSize As Double, ByVal ShrinkageCoefficient As Double, ByVal ShrinkageStartAge As Double, ByVal UseSeries As Long, ByVal NumberSeriesTerms As Long, Optional ByVal Temp As Double = 0) As Long

## Parameters

Name

The name of an existing concrete material property.

ConsiderConcreteAge

If this item is True, time dependence is considered for concrete compressive strength and stiffness (modulus of elasticity).

ConsiderConcreteCreep

If this item is True, time dependence is considered for concrete creep.

ConsiderConcreteShrinkage

If this item is True, time dependence is considered for concrete shrinkage.

CEBFIPsCoefficient

This is the cement type coefficient. This item applies only when ConsiderConcreteAge = True.

RelativeHumidity

This is relative humidity. This item applies only when ConsiderConcreteCreep = True or ConsiderConcreteShrinkage = True.

NotionalSize

This is notional size of the member. This item applies only when ConsiderConcreteCreep = True or ConsiderConcreteShrinkage = True.

As defined in Equation 2.1-69 of CEB\_FIP Model Code 1990 the notional size is equal to two times the cross-sectional area of the member divided by the perimeter of the member in contact with the atmosphere.

ShrinkageCoefficient

This is the shrinkage coefficient as defined in Equation 2.1-76 of CEB\_FIP Model Code 1990. This item applies only when ConsiderConcreteShrinkage = True.

ShrinkageStartAge

This is the shrinkage start age in days as used in Section 2.1.6.4.4 of CEB\_FIP Model Code 1990. This item applies only when ConsiderConcreteShrinkage = True.

UseSeries

This is either 0 or 1, indicating the creep integration type.

0 = Full integration

1 = Dirichlet series

This item applies only when ConsiderConcreteCreep = True.

NumberSeriesTerms

This is the number of series terms used when integrating based on a Dirichlet series. This item applies only when ConsiderConcreteCreep = True and UseSeries = 1.

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data applies. The temperature must have been previously defined for the material.

## Remarks

This function sets the time dependent CEB FIP-90 material property data for concrete materials.

The function returns zero if the data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignMatPropConcreteTimeDepCEBFIP90()

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

      ret = SapModel.PropMaterial.SetMaterial("Concrete", MATERIAL_CONCRETE)



   'assign CEB FIP-90 time dependent data

      ret = SapModel.PropMaterial.TimeDep.SetConcreteCEBFIP90("Concrete", True, True, True, 0.2, 40, 4, 4.9, 2, 1, 12)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetConcreteCEBFIP90](GetConcreteCEBFIP90.md)
