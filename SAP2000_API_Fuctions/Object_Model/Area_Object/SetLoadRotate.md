# SetLoadRotate

## Syntax

SapObject.SapModel.AreaObj.SetLoadRotate

## VB6 Procedure

Function SetLoadRotate(ByVal Name As String, ByVal LoadPat As String, ByVal Value As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

Value

This is the angular velocity. [Cyc/T]

## Remarks

This function assigns rotate loads to area objects.

The function returns zero if the loads are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjectRotateLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'assign area object rotate load

      ret = SapModel.AreaObj.SetLoadRotate("ALL", "DEAD", 30, , Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetLoadRotate](GetLoadRotate_%7BArea_Object%7D.md)

[DeleteLoadRotate](DeleteLoadRotate.md)
