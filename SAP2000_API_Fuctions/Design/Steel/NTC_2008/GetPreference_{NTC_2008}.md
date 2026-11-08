# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.Italian\_NTC\_2008.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

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

3 = K factor method

   1 = Method A

   2 = Method B (default)

   3 = Both

4 = Framing type

1 = DCH-MRF (Default)

2 = DCL-MRF

3 = DCH-CBF

4 = DCL-CBF

5 = DCH-EBF

6 = DCL-EBF

7 = InvPendulum

8 = NonDissipative

5 = GammaM0

Default = 1.05, Value > 0

6 = GammaM1

Default = 1.05, Value > 0

7 = GammaM2

      Default = 1.25, Value > 0

8 = Consider deflection

Default = No = 1, Yes = 2

9 = DL deflection limit, L/Value

Default = 120, Value > 0

10 = SDL + LL deflection limit, L/Value

  Default = 120, Value > 0

11 = LL deflection limit, L/Value

  Default = 360, Value > 0

12 = Total deflection limit, L/Value

  Default = 240, Value > 0

13 = Total camber limit, L/Value

  Default = 240, Value > 0

14 = Pattern live load factor

  Default = 0, Value >= 0

15 = Demand/capacity ratio limit

  Default = 0.95, Value > 0

16 = Multi-response case design

  1 = Envelopes (Default)

  2 = Step-by-step

  3 = Last step

  4 = Envelopes -- All

  5 = Step-by-step -- All

17 = Behavior Factor, q0

  Default = 4, Value > 0

18 = System Overstrength Factor, W

  Default = 1.0, Value > 0

19 = Consider P-Delta

        Default = No = 1, Yes = 2

20 = Consider Torsion

        Default = No = 1, Yes = 2

21 = Ignore Seismic Code

        Default = No = 1, Yes = 2

22 = Ignore Special Seismic Load

        Default = No = 1, Yes = 2

23 = Is Doubler Plate Plug-Welded

  No = 1, Default = Yes = 2

## Remarks

This function retrieves the value of a steel design preference item.

The function returns zero if the item is successfully retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItemItalian_NTC_2008()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value As Double



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

      ret = SapModel.DesignSteel.SetCode("Italian NTC 2008")



   'get preference item

      ret = SapModel.DesignSteel.Italian_NTC_2008.GetPreference(4, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 20.1.0.

## See Also

SetPreference
