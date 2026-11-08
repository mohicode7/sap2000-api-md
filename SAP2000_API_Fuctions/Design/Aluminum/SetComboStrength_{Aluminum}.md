# SetComboStrength

## Syntax

SapObject.SapModel.DesignAluminum.SetComboStrength

## VB6 Procedure

Function SetComboStrength(ByVal Name As String, ByVal Selected As Boolean) As Long

## Parameters

Name

The name of an existing load combination.

Selected

If this item is True, the specified load combination is selected as a design combination for aluminum strength design. If it is False, the combination is not selected for aluminum strength design.

## Remarks

This function selects or deselects a load combination for aluminum strength design.

The function returns zero if the selection status is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAluminumDesignComboStrength()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim i As Long

      Dim NumberNames As Long

      Dim MyName() As String

      Dim Selected As Boolean



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'add aluminum material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_ALUMINUM, , , MATERIAL_ALUMINUM_SUBTYPE_6061_T6)



   'create new aluminum frame section property

      ret = SapModel.PropFrame.SetISection("AI", Name , 18, 6, 0.5, 0.3, 6, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "AI", "AI")



   'add load pattern

      ret = SapModel.LoadPatterns.Add("LIVE", LTYPE_LIVE)



   'add default aluminum design combos

      ret = SapModel.RespCombo.AddDesignDefaultCombos(False, False, True, False)



   'get combo names

      ret = SapModel.RespCombo.GetNameList(NumberNames, MyName)



   'select combos for aluminum strength design

      For i = 0 To NumberNames - 1

         If i = 0 Then Selected = True Else Selected = False

         ret = SapModel.DesignAluminum.SetComboStrength(MyName(i), Selected)

      Next i



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetComboStrength](GetComboStrength_%7BAluminum%7D.md)
