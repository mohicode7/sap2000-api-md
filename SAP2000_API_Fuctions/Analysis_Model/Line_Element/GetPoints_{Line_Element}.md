# GetPoints

## Syntax

SapObject.SapModel.LineElm.GetPoints

## VB6 Procedure

Function GetPoints(ByVal Name As String, ByRef Point1 As String, ByRef Point2 As String) As Long

## Parameters

Name

The name of a defined line element.

Point1

The name of the point element at the I-End of the specified line element.

Point2

The name of the point element at the J-End of the specified line element.

## Remarks

This function retrieves the names of the point elements at each end of a specified line element.

The function returns zero if the point names are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmPoints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Point1 As String

      Dim Point2 As String



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



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get names of points

      ret = SapModel.LineElm.GetPoints("1-1", Point1, Point2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
