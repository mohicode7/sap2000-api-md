# GetAutoWindCode

## Syntax

SapObject.SapModel.LoadPatterns.GetAutoWindCode

## VB6 Procedure

Function GetAutoWindCode(ByVal Name As String, ByRef CodeName As String) As Long

## Parameters

Name

The name of an existing Wind-type load pattern.

CodeName

This is either blank or the name code used for the auto wind parameters. Blank means no auto wind load is specified for the Wind-type load pattern.

## Remarks

This function retrieves the code name used for auto wind parameters in Wind-type load patterns.

The function returns zero if the code is successfully retrieved; otherwise it returns a nonzero value. An error is returned if the specified load pattern is not a Wind-type load pattern.

## VBA Example



```
Sub GetAutoWindType()

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

      ret = SapModel.LoadPatterns.Add("WIND", LTYPE_WIND)



   'get auto wind code

      ret = SapModel.LoadPatterns.GetAutoWindCode("WIND", CodeName)



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

[GetAutoSeismicCode](GetAutoSeismicCode.md)

[GetAutoWaveCode](GetAutoWindCode.md)
