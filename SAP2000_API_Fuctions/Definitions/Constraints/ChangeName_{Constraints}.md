# ChangeName

## Syntax

SapObject.SapModel.ConstraintDef.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined constraint.

NewName

The new name for the constraint.

## Remarks

The function returns zero if the new name is successfully applied, otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangeConstraintName()

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



   'change name of contraint

      ret = SapModel.ConstraintDef.ChangeName("Diaph1", "MyConstraint")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
