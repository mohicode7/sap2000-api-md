# GetCosine

## Syntax

SapObject.SapModel.Func.FuncTH.GetCosine

## VB6 Procedure

Function GetCosine(ByVal Name As String, ByRef CosineP As Double, ByRef CosineSteps As Long, ByRef CosineCycles As Long, ByRef CosineAmp As Double) As Long

## Parameters

Name

The name of a cosine-type time history function.

CosineP

The period of the cosine function. [s]

CosineSteps

The number of steps in the cosine function. This item can not be less than 8.

CosineCycles

The number of cycles in the cosine function.

CosineAmp

The amplitude of the cosine function.

## Remarks

This function retrieves the definition of a cosine-type time history function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTHFuncCosine()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim CosineP As Double

      Dim CosineSteps As Long

      Dim CosineCycles As Long

      Dim CosineAmp As Double



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



   'add cosine TH function

      ret = SapModel.Func.FuncTH.SetCosine("TH-1", 1, 16, 4, 1.25)



   'get cosine TH function

      ret = SapModel.Func.FuncTH.GetCosine("TH-1", CosineP, CosineSteps, CosineCycles, CosineAmp)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetCosine](SetCosine.md)
