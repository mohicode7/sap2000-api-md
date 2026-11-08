# ModifyUndeformedGeometryModeShape

## Syntax

SapObject.SapModel.Analyze.ModifyUnDeformedGeometryModeShape

## VB6 Procedure

Function ModifyUndeformedGeometryModeShape(ByVal CaseName As String, ByVal Mode As Integer, ByVal MaxDisp As Double, ByVal Direction As Integer, Optional ByVal Original As Boolean = False) As Long

## Parameters

CaseName

The name of a modal or buckling load case.

Mode

The mode shape to consider.

MaxDisp

The maximum displacement to which the mode shape will be scaled.

Direction

The direction in which to apply the geometry modification.

Original

If this item is True, all other input items in this function are ignored and the original undeformed geometry data is reinstated.

## Remarks

This function modifies the undeformed geometry based on the shape of a specified mode from a specified modal or buckling load case.

The function returns zero if it is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub ModifyUndeformedGeometryModeShapeExample()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

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

      ret = SapModel.File.NewBeam(2, 288, False)



   'assign point object restraints

      Redim Value(5)

      For i = 0 to 5

         Value(i) = True

      Next i

      ret = SapModel.PointObj.SetRestraint("1", Value)



   'run model

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'modify undeformed geometry

      ret = SapModel.Analyze.ModifyUndeformedGeometryModeShape("MODAL", 1, 1.0, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.2.

## See Also
