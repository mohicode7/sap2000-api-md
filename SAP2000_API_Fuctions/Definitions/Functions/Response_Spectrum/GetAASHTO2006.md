# GetAASHTO2006

## Syntax

SapObject.SapModel.Func.FuncRS.GetAASHTO2006

## VB6 Procedure

Function GetAASHTO2006(ByVal Name As String, ByRef AASHTO2006A As Double, ByRef AASHTO2006SoilProfileType As Long, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of an AASHTO2006 response spectrum function.

AASHTO2006A

The acceleration coefficient, A.

AASHTO2006SoilProfileType

This is 1, 2, 3 or 4, indicating the soil profile type.

1 = I

2 = II

3 = III

4 = IV

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of an AASHTO2006 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncAASHTO2006()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim AASHTO2006A As Double

      Dim AASHTO2006SoilProfileType As Long

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



   'add AASHTO2006 RS function

      ret = SapModel.Func.FuncRS.SetAASHTO2006("RS-1", 0.3, 4, 0.04)



   'get AASHTO2006 RS function

      ret = SapModel.Func.FuncRS.GetAASHTO2006("RS-1", AASHTO2006A, AASHTO2006SoilProfileType, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetAASHTO2006](SetAASHTO2006.md)
