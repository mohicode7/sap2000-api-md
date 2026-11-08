# SetPreference (Concrete Shell Design)

## Syntax

SapObject.SapModel.DesignConcreteShell.Eurocode\_2\_2004.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 9, inclusive, indicating the preference item considered.

1 = Country

2 = Gamma steel

3 = Gamma concrete

4 = AlphaCC

5 = AlphaCT

6 = Crack condition

7 = Shear design method

8 = Cotangent of the angle of concrete compressive strut

9 = Add tensile force due to shear

Value

The value of the considered preference item.

1 = Country

      1 = CEN Default

      2 = United Kingdom

      3 = Slovenia

      5 = Norway

      6 = Singapore

      7 = Sweden

      8 = Finland

      9 = Denmark

    10 = Portugal

    11 = Germany

    12 = Poland

    13 = Ireland

2 = Gamma steel

Value > 0

3 = Gamma concrete

Value > 0

4 = AlphaCC

  Value > 0

5 = AlphaCT

  Value > 0

6 = Crack condition

      1 = Program Determined - the program will perform the calculation according to EN 1992-2:2005 Annex LL (107) to check whether the shell elements are uncracked or cracked

2 = Cracked – the program will assume that the shell elements are cracked without performing the calculation according to EN 1992-2:2005 Annex LL (107)

7 = Shear design method

      1 = Method 1 – considers increasing longitudinal reinforcement to increase concrete shear capacity up to the allowable limit. If insufficient, shear reinforcement will be added

2 = Method 2 – determines required shear reinforcement without considering to increase longitudinal reinforcement

8 = Cotangent of the angle of concrete compressive strut

1 <= Value <= 3.0 for Germany

1 <= Value <= 2.0 for Poland

1 <= Value <= 2.5 for other countries

<a id="_Hlk154131672"></a>9 = Add tensile force due to shear

      0 = No

      Any Other Value = Yes

## Remarks

This function sets the value of a concrete shell design preference item.

The function returns zero if the item is successfully set; otherwise, it returns a nonzero value.

## VBA Example



```
Sub SetConcreteShellDesignPreferenceItemEurocode_2_2004()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'create blank model

             ret = SapModel.File.NewBlank()



       'set concrete design code

          ret = SapModel.DesignConcreteShell.SetCode("Eurocode 2-2004")



       'set preference item

          ret = SapModel.DesignConcreteShell.Eurocode_2_2004.SetPreference(2, 1.1)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 24.0.0

Added items 5 through 8 in v24.1.0.

Added item 9 in v25.2.0.

## See Also

[GetPreference (Concrete Shell EC-2)](GetPreference_%7BConcrete_Shell_Design_Eurocode_2-2004%7D.md)
