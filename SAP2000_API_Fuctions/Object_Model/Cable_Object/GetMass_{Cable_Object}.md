# GetMass

## Syntax

SapObject.SapModel.CableObj.GetMass

## VB6 Procedure

Function GetMass(ByVal Name As String, ByRef MassOverL As Double) As Long

## Parameters

Name

The name of an existing cable object.

MassOverL

The mass per unit length assigned to the cable object. [M/L]

## Remarks

This function retrieves the mass per unit length assignment for cable objects.

The function returns zero if the mass assignment is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim MassOverL As Double



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



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'assign cable mass

      ret = SapModel.CableObj.SetMass("ALL", .0001, False, Group)



   'get cable mass assignment

      ret = SapModel.CableObj.GetMass(Name, MassOverL)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetMass](SetMass_%7BCable_Object%7D.md)

[DeleteMass](DeleteMass_%7BCable_Object%7D.md)
