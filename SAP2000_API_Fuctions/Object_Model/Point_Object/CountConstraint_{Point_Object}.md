# CountConstraint

## Syntax

SapObject.SapModel.PointObj.CountConstraint

## VB6 Procedure

Function CountConstraint(ByRef Count As Long, Optional ByVal Name As String = "") As Long

## Parameters

Count

The number of counted constraints.

Name

This optional item is the name of an existing point object.

## Remarks

If the Name item is provided, the Count item returns the total number of constraint assignments made to the specified point object. If the Name item is not specified, or is specified as an empty string, the Count item returns the total number of constraint assignments to all point objects in the model. If the Name item is specified but it is not recognized by the program as a valid point object, an error is returned.

This function returns zero if the count is successfully completed, otherwise it returns a nonzero value.

## VBA Example



```
Sub CountConstraintAssignments()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim Count as Long



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



   'add constraint definition

      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1")



   'make constraint assignment

      ret = SapModel.PointObj.SetConstraint("3", "Diaph1")



   'get number of constraint assignments

      ret = SapModel.PointObj.CountConstraint(Count)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetConstraint](GetConstraint_%7BPoint_Object%7D.md)

[SetConstraint](SetConstraint.md)

[DeleteConstraint](DeleteConstraint.md)
