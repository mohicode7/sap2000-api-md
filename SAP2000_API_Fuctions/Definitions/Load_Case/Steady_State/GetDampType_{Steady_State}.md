# GetDampType

## Syntax

SapObject.SapModel.LoadCases.SteadyState.GetDampType

## VB6 Procedure

Function GetDampType(ByVal Name As String, ByRef DampType As Long) As Long

## Parameters

Name

The name of an existing steady state load case.

DampType

This is 1 or 2, indicating the hysteretic damping type for the load case.

1 = Constant hysteretic damping for all frequencies

2 = Interpolated hysteretic damping by frequency

## Remarks

This function retrieves the hysteretic damping type for the specified load case.

The function returns zero if the type is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseSteadyStateDampType()

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



   'add response spectrum load case

      ret = SapModel.LoadCases.SteadyState.SetCase("LCASE1")



   'get damping type

      ret = SapModel.LoadCases.SteadyState.GetDampType("LCASE1", DampType)



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

[GetDampConstant](SetDampConstant_%7BSteady_State%7D.md)

[GetDampInterpolated](SetDampInterpolated_%7BSteady_State%7D.md)
