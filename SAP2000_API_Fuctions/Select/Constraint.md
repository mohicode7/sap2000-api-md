# Constraint

## Syntax

SapObject.SapModel.SelectObj.Constraint

## VB6 Procedure

Function Constraint(ByVal Name As String, Optional ByVal DeSelect As Boolean = False) As Long

## Parameters

Name

The name of an existing joint constraint.

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

## Remarks

This function selects or deselects all point objects to which the specified constraint has been assigned.

The function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example



```
Sub SelectConstraint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim i As Long



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



   'define a new constraint

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1")



   'define new constraint assignments

      For i = 4 To 16 Step 4

         ret = SapModel.PointObj.SetConstraint(Format(i), "Diaph1")

      Next i



   'select constraint

      ret = SapModel.SelectObj.Constraint("Diaph1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
