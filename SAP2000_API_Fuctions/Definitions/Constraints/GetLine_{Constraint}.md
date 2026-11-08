# GetLine

## Syntax

SapObject.SapModel.ConstraintDef.GetLine

## VB6 Procedure

Function GetLine(ByVal Name As String, ByRef Value() As Boolean, ByRef CSys As String) As Long

## Parameters

Name

The name of an existing constraint.

Value

Value is an array of six booleans that indicate which joint degrees of freedom are included in the constraint. In order, the degrees of freedom addressed in the array are UX, UY, UZ, RX, RY and RZ.

CSys

The name of the coordinate system in which the constraint is defined.

## Remarks

The function returns the definition for the specified constraint.

The function returns zero if the constraint data is successfully obtained, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineConstraint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim i as long

      Dim ret As Long

      Dim Value() As Boolean

      Dim CSys as String

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



      'define new contraint

      redim Value(5)

      for i = 0 To 5

         Value(i) = True

      Next i

      ret = SapModel.ConstraintDef.SetLine("Line1", Value)



   'get constraint data

      ret = SapModel.ConstraintDef.GetConstraintType("Line1", ConstraintType)

      If ConstraintType = CONSTRAINT_LINE Then

         ret = SapModel.ConstraintDef.GetLine("Line1", Value, CSys)

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

[SetLine](SetLine_%7BConstraints%7D.md)
