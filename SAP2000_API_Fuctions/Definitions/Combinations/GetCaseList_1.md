# GetCaseList\_1

## Syntax

SapObject.SapModel.RespCombo.GetCaseList\_1

## VB6 Procedure

Function GetCaseList\_1(ByVal Name As String, ByRef NumberItems As Long, ByRef CType() As eCNameType, ByRef CName() As String, ByRef ModeNumber() As Long, ByRef SF() As Double) As Long

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

This is an array of integers respresenting the mode numbers for any items where CType=0 and the load case type is either modal or buckling.

ModeNumber

This is an array of integers representing the mode numbers for any items where CType=0 and the load case type is either modal or buckling.

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
Dim ModelNumber() As Long

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

    ret = SapModel.RespCombo.SetCaseList_1("COMB1", eCNameType_LoadCase, "DEAD", 1, 1.4)



    'get all cases and combos included in combo COMB1

    ret = SapModel.RespCombo.GetCaseList_1("COMB1", NumberItems, CType, CName, ModeNumber, SF)



    'close Sap2000

    SapObject.ApplicationExit False

    Set SapModel = Nothing

    Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 22.0.0.

This function supersedes [GetCaseList](GetCaseList.md).

## See Also

[SetCaseList\_1](SetCaseList_1.md)
