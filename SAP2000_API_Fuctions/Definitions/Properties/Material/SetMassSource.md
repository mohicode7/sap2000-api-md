# SetMassSource

## Syntax

SapObject.SapModel.PropMaterial.SetMassSource

## VB6 Procedure

Function SetMassSource(ByVal MyOption As Long, ByVal NumberLoads As Long, ByRef LoadPat() As String, ByRef sf() As Double) As Long

## Parameters

MyOption

This is 1, 2 or 3, indicating the mass source option.

1 = From element self mass and additional masses

2 = From loads

3 = From element self mass and additional masses and loads

NumberLoads

The number of load patterns from which mass is obtained. This item applies only when MyOption is 2 or 3.

LoadPat

This is an array of the names of the load patterns from which mass is obtained. This item applies only when MyOption is 2 or 3.

sf

This is an array of load patterns multipliers used to calculate the mass. This item applies only when MyOption is 2 or 3.

## Remarks

This function sets the mass source for the model.

The function returns zero if the mass source is successfully set; otherwise it returns a nonzero value.

If either the MassFromElements or MassFromMasses parameter is specified as True, both values will be set to true in the model. Both of these items are currently controlled with a single option within the user interface.

## VBA Example



```
Sub AssignMassSource()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyLoadPat() As String

      Dim MySF() As Double



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



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("SDL", LTYPE_SUPERDEAD)



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("LIVE", LTYPE_LIVE)



   'set mass source

      ReDim MyLoadPat(2)

      ReDim MySF(2)

      MyLoadPat(0) = "DEAD"

      MyLoadPat(1) = "SDL"

      MyLoadPat(2) = "LIVE"

      MySF(0) = 1

      MySF(1) = 0.2

      MySF(2) = 0.25

      ret = SapModel.PropMaterial.SetMassSource(3, 3, MyLoadPat, MySF)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetMassSource](GetMassSource.md)
