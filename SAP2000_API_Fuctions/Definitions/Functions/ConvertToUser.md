# ConvertToUser

## Syntax

SapObject.SapModel.Func.ConvertToUser

## VB6 Procedure

Function ConvertToUser(ByVal Name As String) As Long

## Parameters

Name

The name of an existing function that is not a user defined function.

## Remarks

This function converts an existing function to a user defined function.

The function returns zero if the function definition is successfully converted; otherwise it returns a nonzero value. An error is returned if the specified function is already a user defined function.

## VBA Example



```
Sub ConvertFuncToUser()

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



   'add sine TH function

      ret = SapModel.Func.FuncTH.SetSine("TH-1", 1, 16, 4, 1.25)



   'convert to user function

      ret = SapModel.Func.ConvertToUser("TH-1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
