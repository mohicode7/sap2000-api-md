# SetCaseList (Note:  Newer function available)

## Syntax

SapObject.SapModel.RespCombo.SetCaseList

## VB6 Procedure

Function SetCaseList(ByVal Name As String, ByRef CNameType As eCNameType, ByVal CName As String, ByVal SF As Double) As Long

## Parameters

Name

The name of an existing load combination.

CNameType

This is one of the following items in the eCNameType enumeration:

LoadCase = 0

LoadCombo = 1

This item indicates if the CName item is a load case (LoadCase) or a load combination (LoadCombo).

CName

The name of the load case or load combination to be added to or modified in the combination specified by the Name item. If the load case or combination already exists in the combination specified by the Name item, the scale factor is modified as indicated by the SF item for that load case or combination. If the analysis case or combination does not exist in the combination specified by the Name item, it is added.

SF

The scale factor multiplying the case or combination indicated by the CName item.

## Remarks

This function adds or modifies one load case or response combination in the list of cases included in the load combination specified by the Name item.

The function returns zero if the item is successfully added or modified, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddCaseToCombo()

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



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

This function is obsolete and has been superseded by SetCase\_1 as of v22.0.0. This function is maintained for backwards compatibility.

## See Also

[GetCaseList](GetCaseList.md)

[SetCaseList\_1](SetCaseList_1.md)
