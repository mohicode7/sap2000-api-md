# GetNumberModes

## Syntax

SapObject.SapModel.LoadCases.ModalRitz.GetNumberModes

## VB6 Procedure

Function GetNumberModes(ByVal Name As String, ByRef MaxModes As Long, ByRef MinModes As Long) As Long

## Parameters

Name

The name of an existing modal ritz load case.

MaxModes

The maximum number of modes requested.

MinModes

The minimum number of modes requested.

## Remarks

This function retrieves the number of modes requested for the specified load case.

The function returns zero if the number of modes is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseModalRitzNumberModes()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MaxModes As Long

      Dim MinModes As Long



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



   'add modal ritz load case

      ret = SapModel.LoadCases.ModalRitz.SetCase("LCASE1")



   'get number of modes requested

      ret = SapModel.LoadCases.ModalRitz.GetNumberModes("LCASE1", MaxModes, MinModes)



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

[SetNumberModes](SetNumberModes_%7BModal_Ritz%7D.md)
