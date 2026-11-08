# GetOverwrite (Note:  Deprecated, Newer Function Available)

## Syntax

SapObject.SapModel.DesignSteel.Eurocode\_3\_2005.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item
As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a steel frame design
procedure.

Item

This is an integer between 1 and 54, inclusive, indicating
the overwrite item considered.

1 = Framing type

2 = Consider deflection

3 = Deflection check type

4 = DL deflection
limit, L/Value

5 = SDL + LL deflection
limit, L/Value

6 = LL deflection limit,
L/Value

7 = Total load deflection
limit, L/Value

8 = Total camber limit,
L/Value

9 = DL deflection limit,
absolute

10 = SDL + LL deflection
limit, absolute

11 = LL deflection limit,
absolute

12 = Total load deflection
limit, absolute

13 = Total camber limit,
absolute

14 = Specified camber

15 = Net area to total
area ratio

16 = Live load reduction
factor

17 = Unbraced length ratio,
y-y

18 = Unbraced length ratio,
z-z

19 = Effective length
factor, K2y

20 = Effective length
factor, K2z

21 = Moment coefficient,
kyy

22 = Moment coefficient,
kzz

23 = Bending coefficient,
C1

24
= Moment coefficient, kzy

25
= Moment coefficient, kyz

26 = Yield stress, Fy

27 = Compressive capacity,
Nc.Rd

28 = Tensile capacity,
Nt.Rd

29 = Bending capacity
about y-y axis,Mcy.Rd

30 = Bending capacity
about z-z axis Mcz.Rd

31 = Buckling resistance
moment, Mb.Rd

32 = Shear capacity along
z-z axis, Vz.Rd

33 = Shear capacity along
y-y axis, Vy.Rd

34 = Demand/capacity ratio
limit

35 = Section class

36 = Column buckling curve,
y-y

37 = Column buckling curve,
z-z

38 = Buckling curve for LTB

39 = System overstrength factor,
Omega

40 = Is rolled section

41 = Unbraced length ratio,
LTB

42 = Effective length factor
braced, K1y

43 = Effective length factor
braced, K1z

44 = Effective length factor,
K LTB

45 = Material overstrength
factor, GammaOV

46 = Warping constant, Iw

47 = Elastic torsional buckling
force, Ncr T

48 = Elastic torsional-flexural
buckling force, Ncr TF

49
= Bending coefficient, C2

50 = Bending coefficient,
C3

51 = Warping coefficient,
kw

52 = Coordinate of load application,
za

53
= Shear center coordinate, zs

54
= Elastic critical moment for lateral-torsional buckling, Mcr

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = DCH MRF (Ductility
Class High – Moment Frame)

2 = DCM MRF (Ductility
Class Medium – Moment Frame)

3 = DCL MRF (Ductility
Class Low – Moment Frame)

4 = DCH CBF (Ductility
Class High – Concentrically Braced Frame)

5 = DCM CBF (Ductility
Class Medium – Concentrically Braced Frame)

6 = DCL CBF (Ductility
Class Low – Concentrically Braced Frame)

7 = DCH EBF (Ductility
Class High – Eccentrically Braced Frame)

8 = DCM EBF (Ductility
Class Medium – Eccentrically Braced Frame)

9 = DCL EBF (Ductility
Class Low – Eccentrically Braced Frame)

10 = Inverted pendulum
structure

11 = Secondary

2 = Consider deflection

0 = Program Determined

1 = No

2 = Yes

3 = Deflection check type

0 = Program default

1 = Ratio

2 = Absolute

3 = Both

4 = DL deflection limit,
L/Value

Value >= 0; 0 means
no check for this item.

5 = SDL + LL deflection
limit, L/Value

Value >= 0; 0 means
no check for this item.

6 = LL deflection limit,
L/Value

Value >= 0; 0 means
no check for this item.

7 = Total load deflection
limit, L/Value

Value >= 0; 0 means
no check for this item.

8 = Total camber limit,
L/Value

Value >= 0; 0 means
no check for this item.

9 = DL deflection limit,
absolute

Value >= 0; 0 means
no check for this item. [L]

10 = SDL + LL deflection
limit, absolute

  Value >=
0; 0 means no check for this item. [L]

11 = LL deflection limit,
absolute

  Value >=
0; 0 means no check for this item. [L]

12 = Total load deflection
limit, absolute

  Value >=
0; 0 means no check for this item. [L]

13 = Total camber limit,
absolute

  Value >=
0; 0 means no check for this item. [L]

14 = Specified camber

  Value >=
0. [L]

15 = Net area to total
area ratio

  Value >=
0; 0 means use program default value.

16 = Live load reduction
factor

  Value >=
0; 0 means use program determined value.

17 = Unbraced length ratio,
y-y

  Value >=
0; 0 means use program determined value.

18 = Unbraced length ratio,
z-z

  Value >=
0; 0 means use program determined value.

19 = Effective length
factor sway, K2y

  Value >=
0; 0 means use program determined value.

20 = Effective length
factor sway, K2z

  Value >=
0; 0 means use program determined value.

21 = Moment coefficient,
kyy

  Value >=
0; 0 means use program determined value.

22 = Moment coefficient,
kzz

  Value >=
0; 0 means use program determined value.

23 = Bending coefficient,
C1

  Value >=
0; 0 means use program determined value.

24 = Moment coefficient,
kzy

  Value >=
0; 0 means use program determined value.

25 = Moment coefficient,
kyz

  Value >=
0; 0 means use program determined value.

26 = Yield stress, Fy

  Value >=
0; 0 means use program determined value. [F/L2]

27 = Compressive capacity,
Nc.Rd

  Value >=
0; 0 means use program determined value. [F]

28 = Tensile capacity,
Nt.Rd

  Value >=
0; 0 means use program determined value. [F]

29 = Bending capacity
about y-y axis, Mcy.Rd

  Value >=
0; 0 means use program determined value. [FL]

30 = Bending capacity
about z-z axis, Mcz.Rd

  Value >=
0; 0 means use program determined value. [FL]

31 = Buckling resistance
moment, Mb.Rd

  Value >=
0; 0 means use program determined value. [FL]

32 = Shear capacity along
z-z axis, Vz.Rd

  Value >=
0; 0 means use program determined value. [F]

33 = Shear capacity along
y-y axis, Vy.Rd

  Value >=
0; 0 means use program determined value. [F]

34 = Demand/capacity ratio
limit

  Value >= 0;
0 means use program determined value.

35 = Section class

0 = Program default

1 = Class 1

2 = Class 2

3 = Class 3

4 = Class 4

36 = Column buckling curve,
y-y

0 = Program default

1 = a0

2 = a

3 = b

4 = c

5 = d

37 = Column buckling curve,
z-z

0 = Program default

1 = a0

2 = a

3 = b

4 = c

5 = d

38
= Buckling curve for LTB

0 = Program default

1 = a0

2 = a

3 = b

4 = c

5 = d

39 = System overstrength factor,
Omega

Value >= 0; 0 means use
program determined value.

40 = Is rolled section

0 = Program Determined

1 = No

2 = Yes

41 = Unbraced length ratio,
LTB

Value >= 0; 0 means use
program determined value.

42 = Effective length factor
braced, K1y

Value >= 0; 0 means use
program determined value.

43 = Effective length factor
braced, K1z

Value >= 0; 0 means use
program determined value.

44 = Effective length factor,
K LTB

Value >= 0; 0 means use
program determined value.

45 = Material overstrength
factor, GammaOV

Value >= 0; 0 means use
program determined value.

46 = Warping constant, Iw

Value >= 0; 0 means use
program determined value. [L6]

47 = Elastic torsional buckling
force, Ncr T

Value >= 0; 0 means use
program determined value. [F]

48 = Elastic torsional-flexural
buckling force, Ncr TF

Value >= 0; 0 means use
program determined value. [F]

49
= Bending coefficient, C2

Value >= 0; 0 means
use program determined value.

50 = Bending coefficient,
C3

Value >= 0; 0 means
use program determined value.

51 = Warping coefficient,
kw (used in Mcr calculation)

0.5 =<Value =<
1; 0 means use program determined value which is defaulted to 1.0.

52 = Coordinate of load application, za
(used in Mcr calculation)

53
= Shear center coordinate, zs (used in Mcr calculation)

54
= Elastic critical moment for lateral-torsional buckling, Mcr

Value >=
0; 0 means use program determined value. [FL]

ProgDet

If this item is True, the specified value is program
determined.

## Remarks

This function retrieves the value of a steel design
overwrite item.

The function returns zero if the item is successfully
retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignOverwriteItemEurocode_3_2005()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value As Double

      Dim ProgDet As Boolean



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



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start steel design

      ret = SapModel.DesignSteel.StartDesign



   'get overwrite item

      ret = SapModel.DesignSteel.Eurocode_3_2005.GetOverwrite("8", 
 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.0.0.

Updated the list of items in v18.0.0.

Added items 46 – 48 in v18.2.0.

Added items 49 - 52 in v22.0.0

Included all framing types as in Preference form in
SAP2000 Version 22.1.0

Added items 53 - 54 in v22.0.0

Changed references from major and minor axes to y-y
and z-z in version 23.4.0.

The function is DEPRECATED as of version 25.1. Please
use [Get
Overwrite](../Design/Steel/EN_1993_1_1_2005_%28Formerly_Eurocode_3-2005%29/GetOverwrite_%7BSteel_EN_1993-1-1_2005%7D.md). This topic is maintained for reference.

## See Also

[SetOverwrite](SetOverwrite_%7BSteel_Eurocode_3_2005%7D.md)
