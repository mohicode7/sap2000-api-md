# GetCoupledModelType

## Syntax

SapObject.SapModel.PropMaterial.GetCoupledModelType

## VB6 Procedure

Function GetCoupledModelType(ByVal Name As String, ByRef
MatCoupledType As eMatCoupledType, Optional ByVal Temp As Double = 0.0)
As Long

## Parameters

Name

The name of an existing material property.

MatCoupledType

This is one of the following items in the eMatCoupledType
enumeration.

None
= 1

VonMisesPlasticity
= 2

ModifiedDarwinPecknoldConcrete
= 3

Temp

This item applies only
if the specified material has properties that are temperature dependent.
That is, it applies only if properties are specified for the material
at more than one temperature.

This item is the temperature
at which the specified data is to be retrieved. The temperature must have
been defined previously for the material.

## Remarks

This function retrieves
the nonlinear coupled modeling type for a specified material

The function returns
zero if the assignment data is successfully obtained; otherwise it returns
a nonzero value.

## VBA Example



```
Sub GetMatCoupledTypeData()
   'dimension 
 variables
      Dim 
 SapObject as cOAPI
      Dim 
 SapModel As cSapModel
      Dim 
 ret As Long
      Dim 
 MatCoupledType As eMatCoupledType
 
   'create 
 Sap2000 object
      Set 
 SapObject = CreateObject("CSI.SAP2000.API.SapObject")
 
   'start 
 Sap2000 application
      SapObject.ApplicationStart
 
   'create 
 SapModel object
      Set 
 SapModel = SapObject.SapModel
 
   'initialize 
 model
      ret 
 = SapModel.InitializeNewModel
 
   'create 
 model from template
      ret 
 = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)
 
   'initialize 
 new material property
      ret 
 = SapModel.PropMaterial.SetMaterial("Steel", eMatType.eMatType_Steel, 
 -1, "API example test", "Default")
 
   'get 
 coupled material type data
      ret 
 = SapModel.PropMaterial.GetCoupledModelType("Steel", MatCoupledType)
 
   'close 
 Sap2000
      SapObject.ApplicationExit 
 False
      Set 
 SapModel = Nothing
      Set 
 SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 22.1.0

## See Also

[SetCoupledModelType](SetCoupledModelType.md)
