# SetModalCase

## Syntax

SapObject.SapModel.LoadCases.ModHistNonlinear.SetModalCase

## VB6 Procedure

Function SetModalCase(ByVal Name As String, ByVal ModalCase As String) As Long

## Parameters

Name

The name of an existing nonlinear modal history analysis case.

ModalCase

This is the name of an existing modal load case. It specifies the modal load case on which any mode-type load assignments to the specified load case are based.

## Remarks

This function sets the modal case for the specified analysis case.

The function returns zero if the modal case is successfully set; otherwise it returns a nonzero value.

If the specified modal case is not actually a modal case, the program automatically replaces it with the first modal case it can find. If no modal load cases exist, an error is returned.

## VBA Example



```
Sub SetCaseModHistNonlinearModalCase()

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



   'add nonlinear modal history load case

      ret = SapModel.LoadCases.ModHistNonlinear.SetCase("LCASE1")



   'set modal case

      ret = SapModel.LoadCases.ModHistNonlinear.SetModalCase("LCASE1", "MODAL")



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

[GetModalCase](GetModalCase_%7BModal_TH_Nonlinear%7D.md)
