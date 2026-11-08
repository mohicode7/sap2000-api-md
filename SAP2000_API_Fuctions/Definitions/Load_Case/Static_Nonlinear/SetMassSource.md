# SetMassSource

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinear.SetMassSource

## VB6 Procedure

Function SetMassSource(ByVal Name As String, ByVal Source As String) As Long

## Parameters

Name

The name of an existing static nonlinear load case.

Source

This is the name of an existing mass source or a blank string. Blank indicates to use the mass source from the previous load case or the default mass source if the load case starts from zero initial conditions.

## Remarks

This function sets the mass source to be used for the specified load case.

The function returns zero if the mass source is data is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseStaticNonlinearMassSource()

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



   'add static nonlinear load case

      ret = SapModel.LoadCases.StaticNonlinear.SetCase("LCASE1")
   
   'add a new mass source and make it the default mass source
      LoadPat(0) = "DEAD"
      SF(0) = 1.25
      ret = SapModel.SourceMass.SetMassSource("MyMassSource", True, True, True, True, 1, LoadPat, SF)



   'assign a mass source to the static nonlinear load case
      ret = SapModel.LoadCases.StaticNonlinear.SetMassSource("LCASE1", "MyMassSource")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.20.

## See Also

[GetMassSource](GetMassSource.md)
