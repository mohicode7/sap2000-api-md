# GetUser

## Syntax

SapObject.SapModel.Func.FuncTH.GetUser

## VB6 Procedure

Function GetUser(ByVal Name As String, ByRef NumberItems As Long, ByRef MyTime() As Double, ByRef Value() As Double) As Long

## Parameters

Name

The name of a user defined time history function.

NumberItems

The number of frequency and value pairs defined.

MyTime

This is an array that includes the time for each data point. [s]

Value

This is an array that includes the function value for each data point.

## Remarks

This function retrieves the definition of a user defined time history function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTHFuncUser()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Num As Long

      Dim Tmp() As Double

      Dim Val() As Double

      Dim NumberItems As Long

      Dim MyTime() As Double

      Dim Value() As Double



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



   'add user TH function

      NumberItems = 6

      ReDim Tmp(NumberItems - 1)

      ReDim Val(NumberItems - 1)

      Tmp(0) = 0:  Val(0) = 0.1

      Tmp(1) = 1:  Val(1) = 0.02

      Tmp(2) = 2:  Val(2) = -0.06

      Tmp(3) = 3:  Val(3) = -0.02

      Tmp(4) = 4:  Val(4) = 0.05

      Tmp(5) = 5:  Val(5) = 0.02

      ret = SapModel.Func.FuncTH.SetUser("TH-1", NumberItems, Tmp, Val)



   'get user TH function

      ret = SapModel.Func.FuncTH.GetUser("TH-1", NumberItems, MyTime, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetUser](SetUser_%7BTime_History%7D.md)
