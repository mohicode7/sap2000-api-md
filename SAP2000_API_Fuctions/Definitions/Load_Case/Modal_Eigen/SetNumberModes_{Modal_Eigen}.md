# SetNumberModes

## Syntax

SapObject.SapModel.LoadCases.ModalEigen.SetNumberModes

## VB6 Procedure

Function SetNumberModes(ByVal Name As String, ByVal MaxModes As Long, ByVal MinModes As Long) As Long

## Parameters

Name

The name of an existing modal eigen load case.

MaxModes

The maximum number of modes requested.

MinModes

The minimum number of modes requested.

## Remarks

This function sets the number of modes requested for the specified load case.

The function returns zero if the number of modes is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseModalEigenNumberModes()

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



   'add modal eigen load case

      ret = SapModel.LoadCases.ModalEigen.SetCase("LCASE1")



   'set number of modes

      ret = SapModel.LoadCases.ModalEigen.SetNumberModes("LCASE1", 10, 2)



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

[GetNumberModes](GetNumberModes_%7BModal_Eigen%7D.md)
