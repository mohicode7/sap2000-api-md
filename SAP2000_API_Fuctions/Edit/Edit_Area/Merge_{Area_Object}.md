# Merge

## Syntax

SapObject.SapModel.EditArea.Merge

## VB6 Procedure

Function Merge(ByRef NumberAreas As Long, ByRef AreaName() As String) As Long

## Parameters

NumberAreas

The number of originally selected area objects that remain when the merge is successfully completed.

AreaName

This is an array that includes the names of the selected area objects that remain when the merge is successfully completed.

## Remarks

This function merges selected area objects.

The function returns zero if it is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub MergeAreaObjects()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberAreas As Long

      Dim AreaName() As String



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



   'merge area objects

      ret = SapModel.SelectObj.ClearSelection

      ret = SapModel.AreaObj.SetSelected("1", True)

      ret = SapModel.AreaObj.SetSelected("2", True)

      ret = SapModel.AreaObj.SetSelected("4", True)

      ret = SapModel.EditArea.Merge(NumberAreas, AreaName)



   'refresh window

      ret = SapModel.View.RefreshWindow



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
