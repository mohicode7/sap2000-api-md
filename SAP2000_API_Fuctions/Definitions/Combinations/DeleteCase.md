# DeleteCase

## Syntax

SapObject.SapModel.RespCombo.DeleteCase

## VB6 Procedure

Function DeleteCase(ByVal Name As String, ByVal CType As eCNameType, ByVal CName As String) As Long

## Parameters

Name

The name of an existing load combination.

CType

This is one of the following items in the eCNameType enumeration:

LoadCase = 0

LoadCombo = 1

This item indicates whether the CName item is an analysis case (LoadCase) or a load combination (LoadCombo).

CName

The name of the load case or load combination to be deleted from the specified combination.

## Remarks

This function deletes one load case or load combination from the list of cases included in the specified load combination.

The function returns zero if the item is successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteComboCase()

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

      ret = SapModel.RespCombo.Add("COMB1", 0)



   'add load case to combo

      ret = SapModel.RespCombo.SetCaseList("COMB1", eCNameType_LoadCase, "DEAD", 1.4)



   'delete load case from combo

      ret = SapModel.RespCombo.DeleteCase("COMB1", eCNameType_LoadCase, "DEAD")



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

[GetCaseList](GetCaseList.md)

[SetCaseList](SetCaseList.md)
