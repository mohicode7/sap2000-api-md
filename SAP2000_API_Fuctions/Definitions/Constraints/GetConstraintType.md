# GetConstraintType

## Syntax

SapObject.SapModel.ConstraintDef.GetConstraintType

## VB6 Procedure

Function GetConstraintType(ByVal Name As String, ByRef ConstraintType As eConstraintType) As Long

## Parameters

Name

The name of an existing constraint.

ConstraintType

This value is one of the following items in the eConstraintType enumeration.



```
CONSTRAINT_BODY = 1
CONSTRAINT_DIAPHRAGM = 2
CONSTRAINT_PLATE = 3
CONSTRAINT_ROD = 4
CONSTRAINT_BEAM = 5
CONSTRAINT_EQUAL = 6
CONSTRAINT_LOCAL = 7
CONSTRAINT_WELD = 8
CONSTRAINT_LINE = 13
```



## Remarks

The function returns the constraint type for the specified constraint.

The function returns zero if the constraint type is successfully obtained, otherwise it returns a nonzero value.

## VBA Example



```
Sub ConstraintType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim ConstraintType As eConstraintType



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



   'define new constraint

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1", AutoAxis)



   'get constraint type

      ret = SapModel.ConstraintDef.GetConstraintType("Diaph1", ConstraintType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetBeam](GetBeam.md)

[GetBody](GetBody.md)

[GetDiaphragm](GetDiaphragm.md)

[GetEqual](GetEqual.md)

[GetLine](GetLine_%7BConstraint%7D.md)

[GetLocal](GetLocal.md)

[GetPlate](GetPlate_%7BConstraint%7D.md)

[GetRod](GetRod.md)

[GetWeld](GetWeld.md)
