# SetNotionalSize

## Syntax

SapObject.SapModel.PropArea.SetNotionalSize

## VB6 Procedure

Function SetNotionalSize(ByVal Name As String, ByVal stype As String, ByVal Value As Double) As Long

## Parameters

Name

The name of an existing shell-type area section property.

stype

The type to define the notional size of a section. It can be:

"Auto" = Program will determine the notional size based on the average thickness of an area element.

"User" = The notional size is based on the user-defined value.

"None" = Notional size will not be considered. In other words, the time-dependent effect of this section will not be considered.

Value

For stype is "Auto", the Value represents for the scale factor to the program-determined notional size; for **stype** is “User”, the **Value** represents for the user-defined notional size [L]; for **stype** is “None”, the **Value** will not be used and can be set to 1.

## Remarks

This function assigns the method to determine the notional size of an area section for the creep and shrinkage calculations. This function is currently worked for shell type area section.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaPropNotionalSize()
   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long
 Dim stype As String
 Dim Value As Double
 
'create Sap2000 object

   Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



'start Sap2000 application

   SapObject.ApplicationStart



'create SapModel object

   Set SapModel = SapObject.SapModel



'initialize model

   ret = SapModel.InitializeNewModel



'create model from template

   ret = SapModel.File.NewWall(2, 48, 2, 48)



'assign parameters

   stype = “Auto”

   Value = 1.1

   ret = SapModel.PropArea.SetNotionalSize("ASEC1", stype, Value)



'close Sap2000

   SapObject.ApplicationExit False

   Set SapModel = Nothing

   Set SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 16.1.0

## See Also

[GetNotionalSize](GetNotionalSize.md)
