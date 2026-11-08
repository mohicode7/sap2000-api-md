# GetLanesLoaded

## Syntax

SapObject.SapModel.LoadCases.Moving.GetLanesLoaded

## VB6 Procedure

Function GetLanesLoaded(ByVal Name As String, ByVal LoadNumber As Long, ByRef NumberItems As Long, ByRef MyName() As String) As Long

## Parameters

Name

The name of an existing moving load case.

LoadNumber

The load assignment number.

NumberItems

The number of lanes loaded for the specified load assignment number.

MyName

This is an array that includes the name of each lane loaded for the specified load assignment number.

## Remarks

This function retrieves the lanes loaded data for a specified load assignment number in a specified load case.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseMovingLanesLoaded()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyMyClass() As String

      Dim MySF() As Double

      Dim MyMin() As Double

      Dim MyMax() As Double

      Dim MyMyName() As String

      Dim NumberItems As Long

      Dim MyName() As String



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



   'set lanes loaded data for first load assignment

      ReDim MyMyName(0)

      MyMyName(0) = "LANE1"

      ret = SapModel.LoadCases.Moving.SetLanesLoaded("LCASE1", 1, 1, MyMyName)



   'get lanes loaded data for first load assignment

      ret = SapModel.LoadCases.Moving.GetLanesLoaded("LCASE1", 1, NumberItems, MyName)



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

[SetLanesLoaded](GetLanesLoaded.md)
