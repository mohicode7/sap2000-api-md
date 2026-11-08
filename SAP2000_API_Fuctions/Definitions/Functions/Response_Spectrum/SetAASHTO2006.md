# SetAASHTO2006

## Syntax

SapObject.SapModel.Func.FuncRS.SetAASHTO2006

## VB6 Procedure

Function SetAASHTO2006(ByVal Name As String, ByVal AASHTO2006A As Double, ByVal AASHTO2006SoilProfileType As Long, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function then that function is modified, otherwise, a new function is added.

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

This function defines an AASHTO2006 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncAASHTO2006()

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



   'add AASHTO2006 RS function

      ret = SapModel.Func.FuncRS.SetAASHTO2006("RS-1", 0.3, 4, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetAASHTO2006](GetAASHTO2006.md)
