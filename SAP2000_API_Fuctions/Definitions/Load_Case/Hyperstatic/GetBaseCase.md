# GetBaseCase

## Syntax

SapObject.SapModel.LoadCases.HyperStatic.GetBaseCase

## VB6 Procedure

Function GetBaseCase(ByVal Name As String, ByRef BaseCase As String) As Long

## Parameters

Name

The name of an existing hyperstatic load case.

BaseCase

The name of an existing static linear load case that is the base case for the specified hyperstatic load case.

## Remarks

This function retrieves the base case for the specified hyperstatic load case.

The function returns zero if the base case is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseHyperStaticInitialCondition()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim BaseCase As String



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



   'add hyperstatic load case

      ret = SapModel.LoadCases.HyperStatic.SetCase("LCASE1")



   'set base case

      ret = SapModel.LoadCases.HyperStatic.SetBaseCase("LCASE1", "DEAD")



   'get base case

      ret = SapModel.LoadCases.HyperStatic.GetBaseCase("LCASE1", BaseCase)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetBaseCase](SetBaseCase.md)
