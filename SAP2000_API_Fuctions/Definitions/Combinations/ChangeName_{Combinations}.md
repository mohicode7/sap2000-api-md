# ChangeName

## Syntax

SapObject.SapModel.RespCombo.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined load combination.

NewName

The new name for the combination.

## Remarks

The function returns zero if the new name is successfully applied, otherwise it returns a nonzero value.

The new load combination name must be different from all other load combinations and all load cases. If the name is not unique, an error will be returned.

## VBA Example



```
Sub ChangeComboName()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add combo

      ret = SapModel.RespCombo.Add("COMB1", 1)



   'change combo name

      ret = SapModel.RespCombo.ChangeName("COMB1", "MyCombo")



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
