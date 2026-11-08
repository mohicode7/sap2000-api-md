# GetRestraint

## Syntax

SapObject.SapModel.PointElm.GetRestraint

## VB6 Procedure

Function GetRestraint(ByVal Name As String, ByRef Value() As Boolean) As Long

## Parameters

Name

The name of an existing point element.

Value

This is an array of six restraint values.

Value(0) = U1

Value(1) = U2

Value(2) = U3

Value(3) = R1

Value(4) = R2

Value(5) = R3

## Remarks

This function retrieves the restraint assignments for a point element. The restraint assignments are always returned in the point local coordinate system.

The function returns zero if the restraint assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointElmRestraints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value() As Boolean



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



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get point element restraints

      Redim Value(5)

      ret = SapModel.PointElm.GetRestraint("5", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
