# Delete

## Syntax

SapObject.SapModel.ConstraintDef.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing constraint.

## Remarks

The function deletes the specified constraint. All constraint assignments for that constraint are also deleted.

The function returns zero if the constraint is successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteConstraint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'define new contraint

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1", AutoAxis)



   'delete constraint

      ret = SapModel.ConstraintDef.Delete("Diaph1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetBeam](SetBeam.md)

[SetBody](SetBody.md)

[SetDiaphragm](SetDiaphragm.md)

[SetEqual](SetEqual.md)

[SetLine](SetLine_%7BConstraints%7D.md)

[SetLocal](SetLocal.md)

[SetPlate](SetPlate_%7BConstraint%7D.md)

[SetRod](SetRod.md)

[SetWeld](SetWeld.md)
