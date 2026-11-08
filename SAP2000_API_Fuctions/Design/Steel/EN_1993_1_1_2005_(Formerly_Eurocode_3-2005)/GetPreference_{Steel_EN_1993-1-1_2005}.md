# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.EN1993\_1\_1\_2005.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 24, inclusive, indicating the preference item considered.

1 = Country

2 = Combos equation

3 = Reliability class

4 = K factor method

5 = Multi-response case design

6 = Framing type

7 = Behavior factor, q

8 = System overstrength factor, Omega

9 = Consider P-Delta Done

10 = Consider torsion

11 = GammaM0

12 = GammaM1

13 = GammaM2

14 = Ignore seismic code

15 = Ignore special seismic load

16 = Doubler plate is plug-welded

17 = Consider deflection

18 = DL deflection limit, L/Value

19 = SDL + LL deflection limit, L/Value

20 = LL deflection limit, L/Value

21 = Total deflection limit, L/Value

22 = Total camber limit, L/Value

23 = Pattern live load factor

24 = Demand/capacity ratio limit

Value

The value of the considered preference item.

1 = Country

   1 = CEN Default

   2 = United Kingdom

   3 = Slovenia

   4 = Bulgaria

   5 = Norway

   7 = Sweden

   8 = Finland

   9 = Denmark

 10 = Portugal

 11 = Germany

2 = Combos equation

   1 = 1 = Eq. 6.10

   2 = Max of Eqs. 6.10a and 6.10b

3 = Reliability class

   1 = Class 1

   2 = Class 2

   3 = Class 3

4 = K factor method

   1 = Method 1 (Annex A)

   2 = Method 2 (Annex B)

5 = Multi-response case design

1 = Envelopes

2 = Step-by-step

3 = Last step

4 = Envelopes -- All

5 = Step-by-step --  All

6 = Framing type

1 = DCH MRF (Ductility Class High - Moment Frame)

2 = DCM MRF (Ductility Class Medium - Moment Frame)

3 = DCL MRF (Ductility Class Low - Moment Frame)

4 = DCH CBF (Ductility Class High - Concentrically Braced Frame)

5 = DCH CBF (Ductility Class Medium - Concentrically Braced Frame)

6 = DCH CBF (Ductility Class Low - Concentrically Braced Frame)

7 = DCH EBF (Ductility Class High - Eccentrically Braced Frame)

8 = DCM EBF (Ductility Class Medium - Eccentrically Braced Frame)

9 = DCL EBF (Ductility Class Low - Eccentrically Braced Frame)

10 = Inverted pendulum structure

11 = Secondary

7 = Behavior factor, q

Value > 0

8 = System overstrength factor, Omega

Value > 0

9 = Consider P-Delta Done

0 = No

Any other value = Yes

10 = Consider torsion

0 = No

Any other value = Yes

11 = GammaM0

Value > 0

12 = GammaM1

Value > 0

13 = GammaM2

      Value > 0

14 = Ignore seismic code

0 = No

Any other value = Yes

15 = Ignore special seismic load

0 = No

Any other value = Yes

16 = Doubler plate is plug-welded

0 = No

Any other value = Yes

17 = Consider deflection

0 = No

Any other value = Yes

18 = DL deflection limit, L/Value

Value > 0

19 = SDL + LL deflection limit, L/Value

  Value > 0

20 = LL deflection limit, L/Value

  Value > 0

21 = Total deflection limit, L/Value

  Value > 0

22 = Total camber limit, L/Value

  Value > 0

23 = Pattern live load factor

  Value >= 0

24 = Demand/capacity ratio limit

  Value > 0

## Remarks

This function retrieves the value of a steel design preference item.

The function returns zero if the item is successfully retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItemEN1993_1_1_2005()

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

          ret = SapModel.DesignSteel.SetCode("EN 1993-1-1:2005")



       'get preference item

          ret = SapModel.DesignSteel.EN1993_1_1_2005.GetPreference(4, Value)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 25.1.0

## See Also

[SetPreference](SetPreference_%7BSteel_EN_1993-1-1_2005%7D.md)
