# CountPoint

## Syntax

SapObject.SapModel.GDispl.CountPoint

## VB6 Procedure

Function CountPoint(ByVal Name As String, ByRef Count As Long) As Long

## Parameters

Name

The name of an existing generalized displacement.

Count

The number of point objects included in the specified generalized displacement.

## Remarks

This function retrieves the total number of point objects included in a specified generalized displacement.

The function returns zero if the count is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub CountGDisplPoints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SF() As Double

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



   'add generalized displacement

      ret = SapModel.GDispl.Add("GD1", 1)



   'add points to generalized displacement

      ReDim SF(5)

      SF(0) = 0.5

      ret = SapModel.GDispl.SetPoint("GD1", "3", SF)

      ret = SapModel.GDispl.SetPoint("GD1", "7", SF)



   'get number of points in generalized displacement

      ret = SapModel.GDispl.CountPoint("GD1", Count)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetPoint](GetPoint.md)

[SetPoint](SetPoint.md)
