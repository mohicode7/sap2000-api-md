# GetUBC94

## Syntax

SapObject.SapModel.Func.FuncRS.GetUBC94

## VB6 Procedure

Function GetUBC94(ByVal Name As String, ByRef UBC94Z As Double, ByRef UBC94S As Long, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a UBC94 response spectrum function.

UBC94Z

The seismic zone factor, Z.

UBC94S

This is 1, 2 or 3, indicating the soil type.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a UBC94 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncUBC94()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim UBC94Z As Double

      Dim UBC94S As Long

      Dim DampRatio As Double



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



   'add UBC94 RS function

      ret = SapModel.Func.FuncRS.SetUBC94("RS-1", 0.35, 3, 0.04)



   'get UBC94 RS function

      ret = SapModel.Func.FuncRS.GetUBC94("RS-1", UBC94Z, UBC94S, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetUBC94](SetUBC94_%7BRS%7D.md)
