# GetMassSource

## Syntax

SapObject.SapModel.LoadCases.Static Nonlinear.GetMassSource

## VB6 Procedure

Function GetMassSource(ByVal Name As String, ByRef Source As String) As Long

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
Sub GetCaseStaticNonlinearMassSource()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim Source as String

      

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
   
   'get a mass source from the static nonlinear load case
      ret = SapModel.LoadCases.StaticNonlinear.GetMassSource("LCASE1", "Source")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.20.

## See Also

[SetMassSource](SetMassSource.md)
