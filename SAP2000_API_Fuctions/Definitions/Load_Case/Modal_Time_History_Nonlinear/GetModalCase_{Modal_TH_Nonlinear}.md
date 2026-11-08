# GetModalCase

## Syntax

SapObject.SapModel.LoadCases.ModHistNonlinear.GetModalCase

## VB6 Procedure

Function GetModalCase(ByVal Name As String, ByRef ModalCase As String) As Long

## Parameters

Name

The name of an existing nonlinear modal history analysis case.

ModalCase

This is None or the name of an existing modal analysis case. It specifies the modal load case on which any mode-type load assignments to the specified load case are based.

## Remarks

This function retrieves the modal case assigned to the specified load case.

The function returns zero if the modal case is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseModHistNonlinearModalCase()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ModalCase As String



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



   'get modal case

      ret = SapModel.LoadCases.ModHistNonlinear.GetModalCase("LCASE1", ModalCase)



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

[SetModalCase](SetModalCase_%7BModal_TH_Nonlinear%7D.md)
