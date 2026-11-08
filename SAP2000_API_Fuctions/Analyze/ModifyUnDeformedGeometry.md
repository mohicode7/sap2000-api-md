# ModifyUnDeformedGeometry

## Syntax

SapObject.SapModel.Analyze.ModifyUnDeformedGeometry

## VB6 Procedure

Function ModifyUnDeformedGeometry(ByVal CaseName As String, ByVal SF As Double, Optional ByVal Stage As Long = -1, Optional ByVal Original As Boolean = False) As Long

## Parameters

CaseName

The name of the static load case from which displacements are obtained.

SF

The scale factor applied to the displacements.

Stage

This item applies only when the specified load case is a staged construction load case. It is the stage number from which the displacements are obtained. Specifying a -1 for this item means to use the last run stage.

Original

If this item is True, all other input items in this function are ignored and the original undeformed geometry data is reinstated.

## Remarks

This function modifies the undeformed geometry based on displacements obtained from a specified load case.

The function returns zero if it is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub ModifyUnDeformedGeometryExample()

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

      ret = SapModel.Analyze.ModifyUnDeformedGeometry("DEAD", 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also
