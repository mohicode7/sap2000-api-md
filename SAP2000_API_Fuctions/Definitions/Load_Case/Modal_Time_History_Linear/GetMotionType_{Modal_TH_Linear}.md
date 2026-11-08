# GetMotionType

## Syntax

SapObject.SapModel.LoadCases.ModHistLinear.GetMotionType

## VB6 Procedure

Function GetMotionType(ByVal Name As String, ByRef MotionType As Long) As Long

## Parameters

Name

The name of an existing linear modal history analysis case.

MotionType

This is 1 or 2, indicating the time history motion type.

1 = Transient

2 = Periodic

## Remarks

This function retrieves the motion type for the specified load case.

The function returns zero if the motion type is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseModHistLinearMotionType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MotionType As Long



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



   'add linear modal history load case

      ret = SapModel.LoadCases.ModHistLinear.SetCase("LCASE1")



   'get motion type

      ret = SapModel.LoadCases.ModHistLinear.GetMotionType ("LCASE1", MotionType )



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

[SetMotionType](SetMotionType.md)
