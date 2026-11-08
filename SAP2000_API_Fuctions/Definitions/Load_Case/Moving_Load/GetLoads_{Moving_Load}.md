# GetLoads

## Syntax

SapObject.SapModel.LoadCases.Moving.GetLoads

## VB6 Procedure

Function GetLoads(ByVal Name As String, ByRef NumberLoads As Long, ByRef MyClass() As String, ByRef SF() As Double, ByRef Min() As Double, ByRef Max() As Double) As Long

## Parameters

Name

The name of an existing moving load case.

NumberLoads

The number of loads assigned to the specified analysis case.

MyClass

This is an array that includes the vehicle class for each load assigned to the load case.

SF

This is an array that includes the scale factor for each load assigned to the load case.

Min

This is an array that includes the minimum number of lanes loaded for each load assigned to the load case.

Max

This is an array that includes the maximum number of lanes loaded for each load assigned to the load case. This item must be 0, or it must be greater than or equal to Min. If this item is 0, all available lanes are loaded.

## Remarks

This function retrieves the load data for the specified load case.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseMovingLoads()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyMyClass() As String

      Dim MySF() As Double

      Dim MyMin() As Double

      Dim MyMax() As Double

      Dim NumberLoads As Long

      Dim MyClass() As String

      Dim SF() As Double

      Dim Min() As Double

      Dim Max() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open existing model

      SapModel.File.OpenFile("C:\SapAPI\Example 1-030.SDB")



   'add moving load case

      ret = SapModel.LoadCases.Moving.SetCase("LCASE1")



   'set load data

      ReDim MyMyClass(1)

      ReDim MySF(1)

      ReDim MyMin(1)

      ReDim MyMax(1)

      MyMyClass(0) = "VECL1"

      MySF(0) = 1.5

      MyMin(0) = 1

      MyMax(0) = 1

      MyMyClass(1) = "VECL1"

      MySF(1) = 0.8

      MyMin(1) = 0

      MyMax(1) = 0

      ret = SapModel.LoadCases.Moving.SetLoads("LCASE1", 2, MyMyClass, MySF, MyMin, MyMax)



   'get load data

      ret = SapModel.LoadCases.Moving.GetLoads("LCASE1", NumberLoads, MyClass, SF, Max, Min)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SetLoads](SetLoads_%7BMoving_Load%7D.md)
