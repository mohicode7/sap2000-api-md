# GetTriangular

## Syntax

SapObject.SapModel.Func.FuncTH.GetTriangular

## VB6 Procedure

Function GetTriangular(ByVal Name As String, ByRef TriP As Double, ByRef TriCycles As Long, ByRef TriAmp As Double) As Long

## Parameters

Name

The name of a triangular-type time history function.

TriP

The period of the triangular function. [s]

TriCycles

The number of cycles in the function.

TriAmp

The maximum amplitude of the triangular function.

## Remarks

This function retrieves the definition of a triangular-type time history function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTHFuncTriangular()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim TriP As Double

      Dim TriCycles As Long

      Dim TriAmp As Double



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



   'get triangular TH function

      ret = SapModel.Func.FuncTH.GetTriangular("TH-1", TriP, TriCycles, TriAmp)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetTriangular](SetTriangular.md)
