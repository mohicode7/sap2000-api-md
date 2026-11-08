# SetCase

## Syntax

SapObject.SapModel.LoadCases.Buckling.SetCase

## VB6 Procedure

Function SetCase(ByVal Name As String) As Long

## Parameters

Name

The name of an existing or new load case. If this is an existing case then that case is modified, otherwise, a new case is added.

## Remarks

This function initializes a buckling load case. If this function is called for an existing load case, all items for the case are reset to their default value.

The function returns zero if the load case is successfully initialized; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseBuckling()

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



   'add buckling load case

      ret = SapModel.LoadCases.Buckling.SetCase("LCASE1")



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
