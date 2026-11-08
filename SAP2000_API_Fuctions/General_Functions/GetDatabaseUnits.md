# GetDatabaseUnits

## Syntax

SapObject.SapModel.GetDatabaseUnits

## VB6 Procedure

Function GetDatabasetUnits() As eUnits

## Parameters

None

## Remarks

This function returns one of the following items from the eUnits enumeration indicating the database units for the model. All data is internally stored in the model in these units and converted to the present units as needed.

lb\_in\_F = 1

lb\_ft\_F = 2

kip\_in\_F = 3

kip\_ft\_F = 4

kN\_mm\_C = 5

kN\_m\_C = 6

kgf\_mm\_C = 7

kgf\_m\_C = 8

N\_mm\_C = 9

N\_m\_C = 10

Ton\_mm\_C = 11

Ton\_m\_C = 12

kN\_cm\_C = 13

kgf\_cm\_C = 14

N\_cm\_C = 15

Ton\_cm\_C = 16

## VBA Example



```
Sub GetUnitsDatabase()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyUnits As eUnits



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'start a new template model

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'get database units

      MyUnits = SapModel.GetDatabaseUnits



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetPresentUnits](SetPresentUnits.md)

[GetPresentUnits](GetPresentUnits.md)
