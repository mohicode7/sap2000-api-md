# GetCaseList (Note:  Newer function available)

## Syntax

SapObject.SapModel.RespCombo.GetCaseList

## VB6 Procedure

Function GetCaseList(ByVal Name As String, ByRef NumberItems As Long, ByRef CType() As eCNameType, ByRef CName() As String, ByRef SF() As Double) As Long

## Parameters

Name

The name of an existing load combination.

NumberItems

The total number of load cases and load combinations included in the load combination specified by the Name item.

CType

This is an array of one of the following items in the eCNameType enumeration:

LoadCase = 0

LoadCombo = 1

This item indicates if the associated CName item is a load case (LoadCase) or a load combination (LoadCombo).

CName

This is an array of the names of the load cases or load combinations included in the load combination specified by the Name item.

SF

The scale factor multiplying the case or combination indicated by the CName item.

## Remarks

This function returns all load cases and response combinations included in the load combination specified by the Name item.

The function returns zero if the data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCasesInCombo()

       'dimension variables

            Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberItems As Long

          Dim CType() As eCNameType

          Dim CName() As String

          Dim SF() As Double



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



       'get all cases and combos included in combo COMB1

          ret = SapModel.RespCombo.GetCaseList("COMB1", NumberItems, CType, CName, SF)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

This function is obsolete and has been superseded by GetCase\_1 as of v22.0.0. This function is maintained for backwards compatibility.

## See Also

[SetCaseList](SetCaseList.md)

[GetCaseList\_1](GetCaseList_1.md)
