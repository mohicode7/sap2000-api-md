# GetUser

## Syntax

SapObject.SapModel.Func.FuncRS.GetUser

## VB6 Procedure

Function GetUser(ByVal Name As String, ByRef NumberItems As Long, ByRef Period() As Double, ByRef Value() As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a user defined response spectrum function.

NumberItems

The number of frequency and value pairs defined.

Period

This is an array that includes the time for each data point. [s]

Value

This is an array that includes the function value for each data point.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a user defined response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncUser()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Num As Long

      Dim Tmp() As Double

      Dim Val() As Double

      Dim NumberItems As Long

      Dim Period() As Double

      Dim Value() As Double

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



   'add user RS function

      NumberItems = 6

      ReDim Tmp(NumberItems - 1)

      ReDim Val(NumberItems - 1)

      Tmp(0) = 0.03:  Val(0) = 0.4

      Tmp(1) = 0.05:  Val(1) = 2.2

      Tmp(2) = 0.80:  Val(2) = 2.2

      Tmp(3) = 1.20:  Val(3) = 1.0

      Tmp(4) = 4.00:  Val(4) = 0.2

      Tmp(5) = 10.0:  Val(5) = 0.05

      ret = SapModel.Func.FuncRS.SetUser("RS-1", NumberItems, Tmp, Val, 0.04)



   'get user RS function

      ret = SapModel.Func.FuncRS.GetUser("RS-1", NumberItems, Period, Value, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetUser](SetUser_%7BRS%7D.md)
