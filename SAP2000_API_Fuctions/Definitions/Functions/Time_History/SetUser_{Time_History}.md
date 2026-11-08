# SetUser

## Syntax

SapObject.SapModel.Func.FuncTH.SetUser

## VB6 Procedure

Function SetUser(ByVal Name As String, ByVal NumberItems As Long, ByRef MyTime() As Double, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

NumberItems

The number of time and value pairs defined.

MyTime

This is an array that includes the time for each data point. [s]

Value

This is an array that includes the function value for each data point.

## Remarks

This function defines a user time history function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetTHFuncUser()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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

      ReDim MyTime(NumberItems - 1)

      ReDim Value(NumberItems - 1)

      MyTime(0) = 0:  Value(0) = 0.1

      MyTime(1) = 1:  Value(1) = 0.02

      MyTime(2) = 2:  Value(2) = -0.06

      MyTime(3) = 3:  Value(3) = -0.02

      MyTime(4) = 4:  Value(4) = 0.05

      MyTime(5) = 5:  Value(5) = 0.02

      ret = SapModel.Func.FuncTH.SetUser("TH-1", NumberItems, MyTime, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetUser](GetUser_%7BTime_History%7D.md)
