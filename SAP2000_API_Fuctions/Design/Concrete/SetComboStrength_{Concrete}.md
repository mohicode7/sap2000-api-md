# SetComboStrength

## Syntax

SapObject.SapModel.DesignConcrete.SetComboStrength

## VB6 Procedure

Function SetComboStrength(ByVal Name As String, ByVal Selected As Boolean) As Long

## Parameters

Name

The name of an existing load combination.

Selected

If this item is True, the specified load combination is selected as a design combination for concrete strength design. If it is False, the combination is not selected for concrete strength design.

## Remarks

This function selects or deselects a load combination for concrete strength design.

The function returns zero if the selection status is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignComboStrength()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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



   'create new concrete frame section properties

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



   'add load pattern

      ret = SapModel.LoadPatterns.Add("LIVE", LTYPE_LIVE)



   'add default concrete design combos

      ret = SapModel.RespCombo.AddDesignDefaultCombos(False, True, False, False)



   'get combo names

      ret = SapModel.RespCombo.GetNameList(NumberNames, MyName)



   'select combos for concrete strength design

      For i = 0 To NumberNames - 1

         If i = 0 Then Selected = True Else Selected = False

         ret = SapModel.DesignConcrete.SetComboStrength(MyName(i), Selected)

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

[GetComboStrength](GetComboStrength_%7BConcrete%7D.md)
