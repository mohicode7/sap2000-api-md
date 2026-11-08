# GetComboDeflection

## Syntax

SapObject.SapModel.DesignAluminum.GetComboDeflection

## VB6 Procedure

Function GetComboDeflection(ByRef NumberItems As Long, ByRef MyName() As String) As Long

## Parameters

NumberItems

The number of load combinations selected as design combinations for aluminum deflection design.

MyName

This is an array that includes the name of each response combination selected as a design combination for aluminum deflection design.

## Remarks

This function retrieves the names of all load combinations selected as design combinations for aluminum deflection design.

The function returns zero if the names are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAluminumDesignComboDeflection()

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

      ret = SapModel.RespCombo.GetNameList(NumberNames, MyName2)



   'select combos for aluminum deflection design

      For i = 0 To NumberNames - 1

         If i = 0 Then Selected = True Else Selected = False

         ret = SapModel.DesignAluminum.SetComboDeflection(MyName2(i), Selected)

      Next i



   'get combos selected for aluminum deflection design

      ret = SapModel.DesignAluminum.GetComboDeflection(NumberItems, MyName)



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

[SetComboDeflection](SetComboDeflection_%7BAluminum%7D.md)
