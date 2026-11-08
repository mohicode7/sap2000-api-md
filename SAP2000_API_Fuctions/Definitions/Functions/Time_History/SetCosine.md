# SetCosine

## Syntax

SapObject.SapModel.Func.FuncTH.SetCosine

## VB6 Procedure

Function SetCosine(ByVal Name As String, ByVal CosineP As Double, ByVal CosineSteps As Long, ByVal CosineCycles As Long, ByVal CosineAmp As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

CosineP

The period of the cosine function. [s]

CosineSteps

The number of steps in the cosine function. This item can not be less than 8.

CosineCycles

The number of cycles in the cosine function.

CosineAmp

The amplitude of the cosine function.

## Remarks

This function defines a cosine-type time history function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetTHFuncCosine()

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



   'add cosine TH function

      ret = SapModel.Func.FuncTH.SetCosine("TH-1", 1, 16, 4, 1.25)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetCosine](GetCosine.md)
