# GetAutoSeismicCode

## Syntax

SapObject.SapModel.LoadPatterns.GetAutoSeismicCode

## VB6 Procedure

Function GetAutoSeismicCode(ByVal Name As String, ByRef CodeName As String) As Long

## Parameters

Name

The name of an existing Quake-type load pattern.

CodeName

This is either blank or the name code used for the auto seismic parameters. Blank means no auto seismic load is specified for the Quake-type load pattern.

## Remarks

This function retrieves the code name used for auto seismic parameters in Quake-type load patterns.

The function returns zero if the code is successfully retrieved; otherwise it returns a nonzero value. An error is returned if the specified load pattern is not a Quake-type load pattern.

## VBA Example



```
Sub GetAutoSeismicType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim CodeName As String



   'create Sap2000 object

      Set SapObject = New Sap2000v16.SapObject



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



   'get auto seismic code

      ret = SapModel.LoadPatterns.GetAutoSeismicCode("EQX", CodeName)



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

[GetAutoWindCode](GetAutoWindCode.md)

[GetAutoWaveCode](GetAutoWaveCode.md)
