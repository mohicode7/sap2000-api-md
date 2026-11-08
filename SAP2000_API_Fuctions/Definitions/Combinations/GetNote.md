# GetNote

## Syntax

SapObject.SapModel.RespCombo.GetNote

## VB6 Procedure

Function GetNote(ByVal Name As String, ByRef Note As String) As Long

## Parameters

Name

The name of an existing load combination.

Note

The user note, if any, included with the specified combination.

## Remarks

This function retrieves the user note for specified response combination. The note may be blank.

The function returns zero if the note is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetComboNote()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Note As String



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



   'add note to combo

      ret = SapModel.RespCombo.SetNote("COMB1", "My combo note")



   'get note for combo

      ret = SapModel.RespCombo.GetNote("COMB1", Note)



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

[SetNote](SetNote.md)
