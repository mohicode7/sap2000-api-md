# GetTypeOAPI

## Syntax

SapObject.SapModel.PropMaterial.GetTypeOAPI

## VB6 Procedure

Function GetTypeOAPI(ByVal Name As String, ByRef MatType
As eMatType, ByRef SymType As Long) As Long

## Parameters

Name

The name of an existing material property.

MatType

This is one of the following items in the eMatType enumeration.

eMatType\_Steel
= 1

eMatType\_Concrete
= 2

eMatType\_NoDesign
= 3

eMatType\_Aluminum
= 4

eMatType\_ColdFormed
= 5

eMatType\_Rebar
= 6

eMatType\_Tendon
= 7

SymType

This is 0, 1, 2 or 3, indicating the material directional
symmetry type.

0 = Isotropic

1 = Orthotropic

2 = Anisotropic

3 = Uniaxial

## Remarks

This function retrieves the material type for the specified
material.

The function returns zero if the type is successfully
retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetMaterialType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatType As eMatType

      Dim SymType As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'get material type

      ret = SapModel.PropMaterial.GetTypeOAPI("4000Psi", 
 MatType, SymType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Updated the documentation of the eMatType enumeration
in v22.1.0

Initial release in version 11.02.

Changed function name to GetTypeOAPI in v17.0.0.

## See Also
