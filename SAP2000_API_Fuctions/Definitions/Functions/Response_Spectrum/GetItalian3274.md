# GetItalian3274

## Syntax

SapObject.SapModel.Func.FuncRS.GetItalian3274

## VB6 Procedure

Function GetItalian3274(ByVal Name As String, ByRef Italag As Double, ByRef ItalSoilType As Long, ByRef Italq As Double, ByRef ItalLevel As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a Italian 3274 response spectrum function.

Italag

The peak ground acceleration.

ItalSoilType

This is 1, 2, 3, 4 or 5, indicating the seismic intensity.

1 = A

2 = B

3 = C

4 = E

5 = D

Italq

The structure factor.

ItalLevel

This is 0, 1, 2, 3, 4, 5, 6 or 7, indicating the spectral level, direction and building type.

0 = SLU/H/Building

1 = SLU/H/Bridge

2 = SLU/V/Building

3 = SLU/V/Bridge

4 = EL/H/Building

5 = EL/H/Bridge

6 = EL/V/Building

7 = EL/V/Bridge

SLU refers to ultimate strength design and EL refers to elastic design. H and V are horizontal and vertical, respectively.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a Italian 3274 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncItalian3274()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Italag As Double

      Dim ItalSoilType As Long

      Dim Italq As Double

      Dim ItalLevel As Double

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



   'add Italian3274 RS function

      ret = SapModel.Func.FuncRS.SetItalian3274("RS-1", 0.3, 2, 1.1, 4, 0.04)



   'get Italian3274 RS function

      ret = SapModel.Func.FuncRS.GetItalian3274("RS-1", Italag, ItalSoilType, Italq, ItalLevel, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetItalian3274](SetItalian3274.md)
