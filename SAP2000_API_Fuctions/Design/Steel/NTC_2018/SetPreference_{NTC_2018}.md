# SetPreference

## Syntax

SapObject.SapModel.DesignSteel.Italian\_NTC\_2018.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal textValue As String, ByVal numericValue As Double) As Long

## Parameters

Item

This is an integer between 3 and 23, inclusive, indicating the preference item considered.

3 = Method Used for Buckling in P-M-M

4 = Framing type

5 = GammaM0

6 = GammeM1

7 = GammaM2

8 = Consider deflection

9 = DL deflection limit, L/Value

10 = SDL + LL deflection limit, L/Value

11 = LL deflection limit, L/Value

12 = Total deflection limit, L/Value

13 = Total camber limit, L/Value

14 = Pattern live load factor

15 = Demand/capacity ratio limit

16 = Multi-Response Case Design

17 = Behavior Factor, q0

18 = System Overstrength Factor, W

19 = Consider P-Delta

20 = Consider Torsion

21 = Ignore Seismic Code

22 = Ignore Special Seismic Load

23 = Is Doubler Plate Plug-Welded

Value

The value of the considered preference item.

3 = K factor method

1 or "Method A"

2 or "Method B" - Default

3 or "Both"

4 = Framing type

1 or "DCH-MRF"

2 or "DCL-MRF"

3 or "DCH-CBF"

4 or "DCL-CBF"

5 or "DCH-EBF"

6 or "DCL-EBF"

7 or "InvPendulum"

8 or "Non Dissipative" - (Default)

5 = GammaM0

Default = 1.05, Value > 0

6 = GammaM1

Default = 1.05, Value > 0

7 = GammaM2

      Default = 1.25, Value > 0

8 = Consider deflection

1 or "No" - Default

2 or "Yes"

9 = DL deflection limit, L/Value

Default = 0, Value > 0

10 = SDL + LL deflection limit, L/Value

  Default = 0, Value > 0

11 = LL deflection limit, L/Value

  Default = 300, Value > 0

12 = Total deflection limit, L/Value

  Default = 0, Value > 0

13 = Total camber limit, L/Value

  Default = 250, Value > 0

14 = Pattern live load factor

  Default = 0, Value >= 0

15 = Demand/capacity ratio limit

  Default = 1.0, Value > 0

16 = Multi-response case design

  1 or "Envelopes" - (Default)

  2 or "Step-by-step"

  3 or "Last step"

  4 or "Envelopes -- All"

  5 or "Step-by-step -- All"

17 = Behavior Factor, q0

  Default = 1, Value > 0

18 = System Overstrength Factor, W

  Default = 1.0, Value > 0

19 = Consider P-Delta

        1 or "No" - Default

        2 or "Yes"

20 = Consider Torsion

        1 or "No" - Default

        2 or "Yes"

21 = Ignore Seismic Code

        1 or "No" - Default

        2 or "Yes"

22 = Ignore Special Seismic Load

        1 or "No" - Default

        2 or "Yes"

23 = Is Doubler Plate Plug-Welded

  1 or "No" - Default

        2 or "Yes"

## Remarks

This function sets the value of a steel design preference item.

The function returns zero if the item is successfully set; otherwise, it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignPreferenceItemItalian_NTC_2018()

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("Italian NTC 2018")



   'set preference item

      ret = SapModel.DesignSteel.Italian_NTC_2018.SetPreference(4, "DCL-MRF", 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v20.1.0.

## See Also

[GetPreference](GetPreference_%7BNTC_2018%7D.md)
