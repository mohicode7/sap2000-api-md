# SetTriangular

## Syntax

SapObject.SapModel.Func.FuncTH.SetTriangular

## VB6 Procedure

Function SetTriangular(ByVal Name As String, ByVal TriP As Double, ByVal TriCycles As Long, ByVal TriAmp As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

TriP

The period of the triangular function. [s]

TriCycles

The number of cycles in the function.

TriAmp

The maximum amplitude of the triangular function.

## Remarks

This function defines a triangular-type time history function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetTHFuncTriangular()

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



   'add triangular TH function

      ret = SapModel.Func.FuncTH.SetTriangular("TH-1", 1, 4, 1.75)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetTriangular](GetTriangular.md)
