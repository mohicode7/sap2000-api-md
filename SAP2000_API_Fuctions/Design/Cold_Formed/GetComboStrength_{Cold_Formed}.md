# GetComboStrength

## Syntax

SapObject.SapModel.DesignColdFormed.GetComboStrength

## VB6 Procedure

Function GetComboStrength(ByRef NumberItems As Long, ByRef MyName() As String) As Long

## Parameters

NumberItems

The number of load combinations selected as design combinations for cold formed strength design.

MyName

This is an array that includes the name of each response combination selected as a design combination for cold formed strength design.

## Remarks

This function retrieves the names of all load combinations selected as design combinations for cold formed strength design.

The function returns zero if the names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetColdFormedDesignComboStrength()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim i As Long

      Dim NumberNames As Long

      Dim MyName2() As String

      Dim Selected As Boolean

      Dim NumberItems As Long

      Dim MyName() As String



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

      ret = SapModel.RespCombo.GetNameList(NumberNames, MyName2)



   'select combos for cold formed strength design

      For i = 0 To NumberNames – 1

         If i = 0 Then Selected = True Else Selected = False

         ret = SapModel.DesignColdFormed.SetComboStrength(MyName2(i), Selected)

      Next i



   'get combos selected for cold formed strength design

      ret = SapModel.DesignColdFormed.GetComboStrength(NumberItems, MyName)



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

[SetComboStrength](SetComboStrength_%7BCold_Formed%7D.md)
