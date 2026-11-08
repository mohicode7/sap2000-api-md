# SetLine

## Syntax

SapObject.SapModel.ConstraintDef.SetLine

## VB6 Procedure

Function SetLine(ByVal Name As String, ByRef Value() As Boolean, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of an existing constraint.

Value

Value is an array of six booleans that indicate which joint degrees of freedom are included in the constraint. In order, the degrees of freedom addressed in the array are UX, UY, UZ, RX, RY and RZ.

CSys

The name of the coordinate system in which the constraint is defined.

## Remarks

This function defines a Line constraint. If the specified name is not used for a constraint, a new constraint is defined using the specified name. If the specified name is already used for another Line constraint, the definition of that constraint is modified. If the specified name is already used for some constraint that is not a Line constraint, an error is returned.

The function returns zero if the constraint data is successfully added or modified, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetLineConstraint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim i as long

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



   'define new constraint

      redim Value(5)

      for i = 0 To 5

         Value(i) = True

      Next i

      ret = SapModel.ConstraintDef.SetLine("Line1", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetLine](GetLine_%7BConstraint%7D.md)
