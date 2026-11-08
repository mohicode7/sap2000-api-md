# GetDirectionalFactors

## Syntax

SapObject.SapModel.LoadCases.Moving.GetDirectionalFactors

## VB6 Procedure

Function GetDirectionalFactors(ByVal Name As String, ByRef Vertical As Double, ByRef Braking As Double, ByRef Centrifugal As Double) As Integer

## Parameters

Name

The name of an existing moving load case.

Vertical

The moving load directional factor for vertical load.

Braking

The moving load directional factor.

Centrifugal

The moving directional factor for centrifugal load.

## Remarks

This function retrieves the directional factors for the specified load case.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetDirectionalFactors()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Vertical As Double

      Dim Braking As Double

      Dim Centrifugal As Double
 
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
   'set directional factors

      ret = SapModel.LoadCases.Moving.SetDirectionalFactors("LCASE1", 1, 0.7, 0.4)

      

   'get directional factors

      ret = SapModel.LoadCases.Moving.GetDirectionalFactors("LCASE1", 1, Vertical, Braking, Centrifugal)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.2.0.

## See Also

SetDirectionalFactors
