# GetMultiLaneSF

## Syntax

SapObject.SapModel.LoadCases.Moving.GetMultiLaneSF

## VB6 Procedure

Function GetMultiLaneSF(ByVal Name As String, ByRef NumberItems As Long, ByRef SF() As Double) As Long

## Parameters

Name

The name of an existing moving load case.

NumberItems

The number of lanes loaded up to which reduction scale factors are provided.

SF

This is an array that includes the reduction scale factor for the number of lanes loaded from 1 up to NumberItems.

## Remarks

This function retrieves the multilane scale factor data for the specified load case.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseMovingMultiLaneSF()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MySF() As Double

      Dim NumberItems As Long

      Dim SF() As Double



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



   'set multilane scale factor data

      ReDim MySF(1)

      MySF(0) = 1.2

      MySF(1) = 0.8

      ret = SapModel.LoadCases.Moving.SetMultiLaneSF("LCASE1", 2, MySF)



   'get multilane scale factor data

      ret = SapModel.LoadCases.Moving.GetMultiLaneSF("LCASE1", NumberItems, SF)



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

[SetMultiLaneSF](SetMultiLaneSF.md)
