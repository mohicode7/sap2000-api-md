# GetHingeUnloading

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.GetHingeUnloading

## VB6 Procedure

Function GetHingeUnloading(ByVal Name As String, ByRef UnloadType As Long) As Long

## Parameters

Name

The name of an existing static nonlinear staged analysis case.

UnloadType

This is 1, 2 or 3, indicating the hinge unloading option selected for the load case.

1 = Unload entire structure

2 = Apply local redistribution

3 = Restart using secant stiffness

## Remarks

This function retrieves the hinge unloading option for the specified load case.

The function returns zero if the option is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticNonlinearStagedHingeUnloading()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim UnloadType As Long



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



   'add static nonlinear staged load case

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetCase("LCASE1")



   'get hinge unloading option

      ret = SapModel.LoadCases.StaticNonlinearStaged.GetHingeUnloading("LCASE1", UnloadType)



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

[SetHingeUnloading](SetHingeUnloading_%7BStatic_Nonlinear_Staged%7D.md)
