# SetUser

## Syntax

SapObject.SapModel.Func.FuncRS.SetUser

## VB6 Procedure

Function SetUser(ByVal Name As String, ByVal NumberItems As Long, ByRef Period() As Double, ByRef Value() As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

NumberItems

The number of period and value pairs defined.

Period

This is an array that includes the period for each data point. [s]

Value

This is an array that includes the function value for each data point.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a user response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncUser()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim Period() As Double

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



   'add user RS function

      NumberItems = 6

      ReDim Period(NumberItems - 1)

      ReDim Value(NumberItems - 1)

      Period(0) = 0.03:  Value(0) = 0.4

      Period(1) = 0.05:  Value(1) = 2.2

      Period(2) = 0.80:  Value(2) = 2.2

      Period(3) = 1.20:  Value(3) = 1.0

      Period(4) = 4.00:  Value(4) = 0.2

      Period(5) = 10.0:  Value(5) = 0.05

      ret = SapModel.Func.FuncRS.SetUser("RS-1", NumberItems, Period, Value, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetUser](GetUser_%7BRS%7D.md)
