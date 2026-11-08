# GetDampType

## Syntax

SapObject.SapModel.LoadCases.ModHistNonlinear.GetDampType

## VB6 Procedure

Function GetDampType(ByVal Name As String, ByRef DampType As Long) As Long

## Parameters

Name

The name of an existing nonlinear modal history analysis case.

DampType

This is 1, 2, 3, 4, 5 or 6, indicating the modal damping type for the load case.

1 = Mass and stiffness proportional damping by direct specification

2 = Mass and stiffness proportional damping by period

3 = Mass and stiffness proportional damping by frequency

4 = Constant damping for all modes

5 = Interpolated damping by period

6 = Interpolated damping by frequency

## Remarks

This function retrieves the modal damping type for the specified load case.

The function returns zero if the type is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseModHistNonlinearDampType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DampType As Long



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



   'get damping type

      ret = SapModel.LoadCases.ModHistNonlinear.GetDampType("LCASE1", DampType)



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

[GetDampConstant](GetDampConstant_%7BModal_TH_Nonlinear%7D.md)

[GetDampInterpolated](GetDampInterpolated_%7BModal_TH_Nonlinear%7D.md)

[GetDampProportional](GetDampProportional_%7BModal_TH_Nonlinear%7D.md)

[GetDampOverrides](GetDampOverrides_%7BModal_TH_Nonlinear%7D.md)
