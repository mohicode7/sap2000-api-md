# SetPoint

## Syntax

SapObject.SapModel.GDispl.SetPoint

## VB6 Procedure

Function SetPoint(ByVal Name As String, ByVal PointName As String, ByRef SF() As Double) As Long

## Parameters

Name

The name of an existing generalized displacement.

PointName

The name of a point object to be included in the generalized displacement definition.

SF

This is an array of six unitless scale factors for the point object displacement degrees of freedom.

SF(0) = U1 scale factor

SF(1) = U2 scale factor

SF(2) = U3 scale factor

SF(3) = R1 scale factor

SF(4) = R2 scale factor

SF(5) = R3 scale factor

## Remarks

This function adds a point object and its scale factors to a generalized displacement definition, or, if the point object already exists in the generalized displacement definition, it modifies the scale factors.

The function returns zero if the data is successfully added or modified, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetGDisplPointData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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



   'add point to generalized displacement

      ReDim SF(5)

      SF(0) = 0.5

      ret = SapModel.GDispl.SetPoint("GD1", "3", SF)



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
