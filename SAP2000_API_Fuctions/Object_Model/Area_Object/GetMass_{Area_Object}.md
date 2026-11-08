# GetMass

## Syntax

SapObject.SapModel.AreaObj.GetMass

## VB6 Procedure

Function GetMass(ByVal Name As String, ByRef MassOverL2 As Double) As Long

## Parameters

Name

The name of an existing area object.

MassOverL2

The mass per unit area assigned to the area object. [M/L2]

## Remarks

This function retrieves the mass per unit area assignment for area objects.

The function returns zero if the mass assignment is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MassOverL2 As Double



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



   'assign area object mass

      ret = SapModel.AreaObj.SetMass("ALL", .0001, False, Group)



   'get area object mass assignment

      ret = SapModel.AreaObj.GetMass("1", MassOverL2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetMass](SetMass_%7BArea_Object%7D.md)

[DeleteMass](DeleteMass_%7BArea_Object%7D.md)
