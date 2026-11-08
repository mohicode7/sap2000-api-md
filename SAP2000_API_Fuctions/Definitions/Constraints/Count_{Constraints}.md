# Count

## Syntax

SapObject.SapModel.ConstraintDef.Count

## VB6 Procedure

Function Count(Optional ByVal ConstraintType As eConstraintType) As Long

## Parameters

ConstraintType

This optional value is one of the following items in the eConstraintType enumeration.



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

If the ConstraintType item is specified,  the function returns the number of defined constraints of the specified type. Otherwise it returns the total number of defined constraints without regard for type.

## VBA Example



```
Sub CountConstraints()

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



   'return number of defined constraints

      ret = SapModel.ConstraintDef.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Modified optional argument ConstraintType to be ByVal in version 12.0.1.

## See Also
