# SetCoupledModelType3D

## Syntax

SapObject.SapModel.PropMaterial.SetCoupledModelType3D

## VB6 Procedure

Function SetCoupledModelType3D(ByVal Name As String, ByVal MatCoupledType As eMatCoupledType, Optional ByVal Temp As Double = 0.0) As Long

## Parameters

Name

The name of an existing concrete material property.

MatCoupledType

This is one of the following items in the eMatCoupledType enumeration.

1 = None

2 = VonMisesPlasticity

3 = ModifiedDarwinPecknoldConcrete

4 = FariaConcreteDamagePlasticity

Temp

This item applies only if the specified material has properties that are temperature dependent. That is, it applies only if properties are specified for the material at more than one temperature.

This item is the temperature at which the specified data is to be retrieved. The temperature must have been defined previously for the material.

## Remarks

This function sets the nonlinear 3D coupled modeling type for a specified concrete material.  
The function returns zero if the assignment data is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetMatCoupledType3DData()

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

          ret = SapModel.PropMaterial.SetMaterial("Concrete", eMatType.eMatType_Concrete, -1, "API example test", "Default")
 
   'set coupled material type data

          ret = SapModel.PropMaterial.SetCoupledModelType3D("Concrete", eMatCoupledType.eMatCoupledType_FariaConcreteDamagePlasticity)
 
   'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub

     
```



## Release Notes

Initial release in version 27.1.0

## See Also

[GetCoupledModelType](GetCoupledModelType.md)

[SetCoupledModelType](SetCoupledModelType.md)

[GetCoupledModelType3D](GetCoupledModelType3D.md)

[SetCoupledModelType3D](GetCoupledModelType3D.md)
