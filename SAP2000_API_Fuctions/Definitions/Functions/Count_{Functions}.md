# Count

## Syntax

SapObject.SapModel.Func.Count

## VB6 Procedure

Function Count(Optional ByVal FuncType As Long = 0) As Long

## Parameters

FuncType

This optional value is one of the following numbers, indicating the type of function for which the count is desired.

0 = All function types

1 = Response spectrum

2 = Time history

3 = Power spectral density

4 = Steady state

## Remarks

This function returns the total number of defined functions in the model of the specified type.

## VBA Example



```
Sub CountFunctions()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'return number of defined functions of all types

      Count = SapModel.Func.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

## See Also
