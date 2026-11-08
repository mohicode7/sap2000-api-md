# DeletePoint

## Syntax

SapObject.SapModel.GDispl.DeletePoint

## VB6 Procedure

Function DeletePoint(ByVal Name As String, ByVal PointName As String) As Long

## Parameters

Name

The name of an existing generalized displacement.

PointName

The name of a point object included in the generalized displacement that is to be deleted.

## Remarks

This function deletes one point object from a generalized displacement definition.

The function returns zero if the point is successfully deleted from the generalized displacement definition, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteGDisplPoint()

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



   'add points to generalized displacement

      ReDim SF(5)

      SF(0) = 0.5

      ret = SapModel.GDispl.SetPoint("GD1", "3", SF)

      ret = SapModel.GDispl.SetPoint("GD1", "7", SF)



   'delete point from generalized displacement

      ret = SapModel.GDispl.DeletePoint("GD1", "3")



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
