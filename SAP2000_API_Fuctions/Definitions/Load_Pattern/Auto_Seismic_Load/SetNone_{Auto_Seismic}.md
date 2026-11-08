# SetNone

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeismic.SetNone

## VB6 Procedure

Function SetNone(ByVal Name As String) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

## Remarks

This function sets the auto seismic loading type for the specified load pattern to None.

The function returns zero if the loading type is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSeismicNone()

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



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("EQX", LTYPE_QUAKE)



   'assign BOCA96 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetBOCA96("EQX", 1, 0.05, 1, 0.035, 0, False, 0, 0, 0.4, 0.4, 1.5, 8)



   'set auto seismic loading type to None

      ret = SapModel.LoadPatterns.AutoSeismic.SetNone("EQX")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also
