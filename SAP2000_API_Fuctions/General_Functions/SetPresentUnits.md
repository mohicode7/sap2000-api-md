# SetPresentUnits

## Syntax

SapObject.SapModel.SetPresentUnits

## VB6 Procedure

Function SetPresentUnits(ByVal Units As eUnits) As Long

## Parameters

Units

One of the following items in the eUnits enumeration:

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

## Remarks

This function returns zero if the units are successfully set and nonzero if they are not set.

## VBA Example



```
Sub SetUnitsPresent()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim FileName As String

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'start a new template model

      ret = SapModel.File.New2DFrame(ConcentricBraced, 3, 124, 3, 200)



   'set present units to KN-m

      ret = SapModel.SetPresentUnits(KN_m_C)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetDatabaseUnits](GetDatabaseUnits.md)

[GetPresentUnits](GetPresentUnits.md)
