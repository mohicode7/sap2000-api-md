# GetComboStrength

## Syntax

SapObject.SapModel.DesignSteel.GetComboStrength

## VB6 Procedure

Function GetComboStrength(ByRef NumberItems As Long, ByRef MyName() As String) As Long

## Parameters

NumberItems

The number of load combinations selected as design combinations for steel strength design.

MyName

This is an array that includes the name of each response combination selected as a design combination for steel strength design.

## Remarks

This function retrieves the names of all load combinations selected as design combinations for steel strength design.

The function returns zero if the names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignComboStrength()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add load pattern

      ret = SapModel.LoadPatterns.Add("LIVE", LTYPE_LIVE)



   'add default steel design combos

      ret = SapModel.RespCombo.AddDesignDefaultCombos(True, False, False, False)



   'get combo names

      ret = SapModel.RespCombo.GetNameList(NumberNames, MyName2)



   'select combos for steel strength design

      For i = 0 To NumberNames - 1

         If i = 0 Then Selected = True Else Selected = False

         ret = SapModel.DesignSteel.SetComboStrength(MyName2(i), Selected)

      Next i



   'get combos selected for steel strength design

      ret = SapModel.DesignSteel.GetComboStrength(NumberItems, MyName)



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

[SetComboStrength](SetComboStrength_%7BSteel%7D.md)
