# SetComboDeflection

## Syntax

SapObject.SapModel.DesignColdFormed.SetComboDeflection

## VB6 Procedure

Function SetComboDeflection(ByVal Name As String, ByVal Selected As Boolean) As Long

## Parameters

Name

The name of an existing load combination.

Selected

If this item is True, the specified load combination is selected as a design combination for cold formed deflection design. If it is False, the combination is not selected for cold formed deflection design.

## Remarks

This function selects or deselects a load combination for cold formed deflection design.

The function returns zero if the selection status is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetColdFormedDesignComboDeflection()

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



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section property

      ret = SapModel.PropFrame.SetColdC("CdC", Name , 9, 3, 0.06, 0.25, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 72, 2, 144, True, "CdC", "CdC")



   'add load pattern

      ret = SapModel.LoadPatterns.Add("LIVE", LTYPE_LIVE)



   'add default cold formed design combos

      ret = SapModel.RespCombo.AddDesignDefaultCombos(False, False, False, True)



   'get combo names

      ret = SapModel.RespCombo.GetNameList(NumberNames, MyName)



   'select combos for cold formed deflection design

      For i = 0 To NumberNames – 1

         If i = 0 Then Selected = True Else Selected = False

         ret = SapModel.DesignColdFormed.SetComboDeflection(MyName(i), Selected)

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

[GetComboDeflection](GetComboDeflection_%7BCold_Formed%7D.md)
