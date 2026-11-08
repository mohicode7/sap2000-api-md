# GetPoint

## Syntax

SapObject.SapModel.GDispl.GetPoint

## VB6 Procedure

Function GetPoint(ByVal Name As String, ByRef NumberItems As Long, ByRef PointName() As String, ByRef U1() As Double, ByRef U2() As Double, ByRef U3() As Double, ByRef R1() As Double, ByRef R2() As Double, ByRef R3() As Double) As Long

## Parameters

Name

The name of an existing generalized displacement.

NumberItems

The number of point objects included in the generalized displacement definition.

PointName

This is an array that includes the name of the point objects included in the generalized displacement definition.

U1, U2, U3, R1, R2, R3

These are arrays that include the unitless scale factors for each of the displacement degrees of freedom of the associated point objects that are included in the generalized displacement definition.

## Remarks

This function retrieves the point objects and their scale factors from a generalized displacement definition.

The function returns zero if the data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetGDisplPointData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim PointName() As String

      Dim U1() As Double

      Dim U2() As Double

      Dim U3() As Double

      Dim R1() As Double

      Dim R2() As Double

      Dim R3() As Double

      Dim SF() As Double



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



   'get point data from generalized displacement

      ret = SapModel.GDispl.GetPoint("GD1", NumberItems, PointName, U1, U2, U3, R1, R2, R3)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetPoint](SetPoint.md)
