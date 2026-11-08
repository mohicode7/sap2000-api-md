# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.Norsok\_N0042013.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 16, inclusive, indicating the preference item considered.

1 = Combos equation

2 = K factor method

3 = Pressure interaction method

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

16 = Multi-response case design

Value

The value of the considered preference item.

1 = Combos equation

      1 = 1 = Eq. 6.10

      2 = Max of Eqs. 6.10a and 6.10b

2 = K factor method

      1 = Method 1 (Annex A)

      2 = Method 2 (Annex B)

3 = Pressure interaction method

      1 = Method A

      2 = Method B

4 = Framing type

1 = Moment Frame

2 = Braced Frame

5 = GammaM0

 Value > 0

6 = GammeM1

 Value > 0

7 = GammeM2

      Value > 0

8 = Consider deflection

0 = No

Any other value = Yes

9 = DL deflection limit, L/Value

Value > 0

10 = SDL + LL deflection limit, L/Value

  Value > 0

11 = LL deflection limit, L/Value

  Value > 0

12 = Total deflection limit, L/Value

  Value > 0

13 = Total camber limit, L/Value

  Value > 0

14 = Pattern live load factor

  Value >= 0

15 = Demand/capacity ratio limit

  Value > 0

16 = Multi-response case design

  1 = Envelopes

  2 = Step-by-step

  3 = Last step

  4 = Envelopes -- All

  5 = Step-by-step -- All

## Remarks

This function retrieves the value of a steel design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItemNorsok_N0042013()

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

      ret = SapModel.DesignSteel.SetCode("Norsok N-004 2013")



   'get preference item

      ret = SapModel.DesignSteel.Norsok_N0042013.GetPreference(4, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 16.1.0.

## See Also

[SetPreference](../Norsok_N-004/SetPreference_%28Steel_Norosk_N-004%29.md)
