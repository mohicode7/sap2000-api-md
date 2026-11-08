# GetNameList

## Syntax

SapObject.SapModel.PropMaterial.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef
MyName() As String, Optional ByVal MatType As eMatType) As Long

## Parameters

NumberNames

The number of material property names retrieved by the
program.

MyName

This is a one-dimensional array of material property
names. The MyName array is created as a dynamic, zero-based, array by
the API user:



```
Dim MyName() 
 as String
```



The array is dimensioned to (NumberNames - 1) inside
the SAP2000 program, filled with the names, and returned to the API user.

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

If no value is input for MatType, names are returned
for all material properties in the model regardless of type.

## Remarks

This function retrieves the names of all defined material
properties of the specified type.

The function returns zero if the names are successfully
retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetMaterialNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberNames As Long

      Dim MyName() As String



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



   'get material property names

      ret = SapModel.PropMaterial.GetNameList(NumberNames, 
 MyName)



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
