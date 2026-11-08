# GetLocal

## Syntax

SapObject.SapModel.ConstraintDef.GetLocal

## VB6 Procedure

Function GetLocal(ByVal Name As String, ByRef Value() As Boolean) As Long

## Parameters

Name

The name of an existing constraint.

Value

Value is an array of six booleans that indicate which joint degrees of freedom are included in the constraint. In order, the degrees of freedom addressed in the array are U1, U2, U3, R1, R2 and R3.

## Remarks

The function returns the definition for the specified constraint.

The function returns zero if the constraint data is successfully obtained, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLocalConstraint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim i as long

      Dim ret As Long

      Dim Value() As Boolean

      Dim UY as Boolean

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

      redim Value(5)

      for i = 0 To 5

         Value(i) = True

      Next i

      ret = SapModel.ConstraintDef.SetLocal("Local1", Value)



   'get constraint data

      ret = SapModel.ConstraintDef.GetConstraintType("Local1", ConstraintType)

      If ConstraintType = CONSTRAINT_LOCAL Then

         ret = SapModel.ConstraintDef.GetLocal("Local1", Value)

         if ret = 0 Then UY = Value(1)

      End If



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetLocal](SetLocal.md)
