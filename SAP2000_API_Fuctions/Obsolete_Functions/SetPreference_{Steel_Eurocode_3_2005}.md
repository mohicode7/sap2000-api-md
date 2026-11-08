# SetPreference (Note:  Deprecated, Newer Function Available)

## Syntax

SapObject.SapModel.DesignSteel.Eurocode\_3\_2005.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value
As Double) As Long

## Parameters

Item

This is an integer between 1 and 16, inclusive, indicating
the preference item considered.

1 = Country

2 = Combos equation

3 = Reliability class

4 = K factor method

5 = Multi-response case
design

6 = Framing type

7 = Behavior factor, q

8 = System overstrength
factor, Omega

9 = Consider P-Delta Done

10 = Consider torsion

11 = GammaM0

12 = GammaM1

13 = GammaM2

14 = Ignore seismic code

15 = Ignore special seismic
load

16 = Doubler plate is
plug-welded

17 = Consider deflection

18 = DL deflection limit,
L/Value

19 = SDL + LL deflection
limit, L/Value

20 = LL deflection limit,
L/Value

21 = Total deflection
limit, L/Value

22 = Total camber limit,
L/Value

23 = Pattern live load
factor

24 = Demand/capacity ratio
limit

Value

The value of the considered preference item.

1 = Country

   1 =
CEN Default

   2 =
United Kingdom

   3 =
Slovenia

   4 =
Bulgaria

   5 =
Norway

   7 =
Sweden

   8 =
Finland

   9 =
Denmark

 10 = Portugal

 11 = Germany

2 = Combos equation

   1 =
1 = Eq. 6.10

   2 =
Max of Eqs. 6.10a and 6.10b

3 = Reliability class

   1 =
Class 1

   2 =
Class 2

   3 =
Class 3

4 = K factor method

   1 =
Method 1 (Annex A)

   2 =
Method 2 (Annex B)

5 = Multi-response case
design

1 = Envelopes

2 = Step-by-step

3 = Last step

4 = Envelopes -- All

5 = Step-by-step --  All

6 = Framing type

1 = DCH MRF (Ductility
Class High - Moment Frame)

2 = DCM MRF (Ductility
Class Medium - Moment Frame)

3 = DCL MRF (Ductility
Class Low - Moment Frame)

4 = DCH CBF (Ductility
Class High - Concentrically Braced Frame)

5 = DCH CBF (Ductility
Class Medium - Concentrically Braced Frame)

6 = DCH CBF (Ductility
Class Low - Concentrically Braced Frame)

7 = DCH EBF (Ductility
Class High - Eccentrically Braced Frame)

8 = DCM EBF (Ductility
Class Medium - Eccentrically Braced Frame)

9 = DCL EBF (Ductility
Class Low - Eccentrically Braced Frame)

10 = Inverted pendulum
structure

11 = Secondary

7 = Behavior factor, q

Value > 0

8 = System overstrength
factor, Omega

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

      Value
> 0

14 = Ignore seismic code

0 = No

Any other value = Yes

15 = Ignore special seismic
load

0 = No

Any other value = Yes

16 = Doubler plate is
plug-welded

0 = No

Any other value = Yes

17 = Consider deflection

0 = No

Any other value = Yes

18 = DL deflection limit,
L/Value

Value > 0

19 = SDL + LL deflection
limit, L/Value

  Value >
0

20 = LL deflection limit,
L/Value

  Value >
0

21 = Total deflection
limit, L/Value

  Value >
0

22 = Total camber limit,
L/Value

  Value >
0

23 = Pattern live load
factor

  Value >=
0

24 = Demand/capacity ratio
limit

  Value >
0

## Remarks

This function sets the value of a steel design preference
item.

The function returns zero if the item is successfully
set; otherwise, it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignPreferenceItemEurocode_3_2005()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("Eurocode 
 3-2005")



   'set preference item

      ret = SapModel.DesignSteel.Eurocode_3_2005.SetPreference(4, 
 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.0.0.

Added Norway as a Country parameter in version 14.1.0.

Added Reliability Class parameter and added Sweden,
Finland, and Denmark as Country parameters in version 14.2.2.

Added Portugal and Germany as Country parameters in
SAP2000 Version 15.0.0.

Changed Time history design item to Multi-response case
design and added additional values in version 15.0.1.

Included all parameters and all framing types as in
Preference form in SAP2000 Version 22.1.0

The function is DEPRECATED as of version 25.1. Please
use [Set
Preferences](../Design/Steel/EN_1993_1_1_2005_%28Formerly_Eurocode_3-2005%29/SetPreference_%7BSteel_EN_1993-1-1_2005%7D.md). This topic is maintained for reference.

## See Also

[GetPreference](GetPreference_%7BSteel_Eurocode_3_2005%7D.md)
