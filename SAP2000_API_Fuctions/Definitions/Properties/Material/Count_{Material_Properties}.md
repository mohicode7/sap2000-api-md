# Count

## Syntax

SapObject.SapModel.PropMaterial.Count

## VB6 Procedure

Function Count(Optional ByVal MatType As eMatType) As
Long

## Parameters

MatType

This optional value is one of the following items in
the eMatType enumeration.

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

If no value is input for MatType, a count is returned
for all material properties in the model regardless of type.

## Remarks

This function returns the total number of defined material
properties in the model. If desired, counts can be returned for all material
properties of a specified type in the model.

## VBA Example



```
Sub CountMaterials()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long



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
 3, 124, 3, 200)



   'return number of defined materials of all types

      Count = SapModel.PropMaterial.Count



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

## See Also
