# GetSpecialRigidDiaphragmList

## Syntax

SapObject.SapModel.ConstraintDef.GetSpecialRigidDiaphragmList

## VB6 Procedure

Function GetSpecialRigidDiaphragmList(ByRef Num As Long, ByRefDiaph() As String) As Long

## Parameters

Num

The number of special rigid diaphragm constraints.

Diaph

This is an array that includes the name of each special rigid diaphragm constraint.

## Remarks

This function retrieves the list of the names of each special rigid diaphragm constraint. A special rigid diaphragm constraint is required for assignment of auto seismic load diaphragm eccentricity overwrites. It is also required for calculation of auto wind loads whose exposure widths are determined from the extents of rigid diaphragms.

A special rigid diaphragm constraint is a constraint with the following features:

1. The constraint type is CONSTRAINT\_DIAPHRAGM = 2.

2. The constraint coordinate system is Global.

3. The constraint axis is Z.

The function returns zero if the name list is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSpecialDiaphragmList()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Num As Long

      Dim Diaph() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'createSapModelobject

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New3DFrame(BeamSlab, 2, 144, 3, 336, 2, 432)



   'define diaphragm constraints

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1", Z)

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph2")

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph3", Z)



   'get special rigid diaphragm name list

      ret = SapModel.ConstraintDef.GetSpecialRigidDiaphragmList(Num, Diaph)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
