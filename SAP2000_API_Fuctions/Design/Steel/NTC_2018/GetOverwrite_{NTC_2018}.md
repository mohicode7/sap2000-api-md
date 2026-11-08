# GetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.Italian\_NTC\_2018.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item
As Long, ByRef textValue As String, ByRef numericValue As Double, ByRef
ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a steel frame design
procedure.

Item

This is an integer between 1 and 55, inclusive, indicating
the overwrite item considered.

1 = Current Design Section

2 = Framing type

3 = Consider deflection

4 = Deflection check type

5 = DL deflection limit,
L/Value

6 = SDL + LL deflection
limit, L/Value

7 = LL deflection limit,
L/Value

8 = Total load deflection
limit, L/Value

9 = Total camber limit,
L/Value

10 = DL deflection limit,
absolute

11 = SDL + LL deflection
limit, absolute

12 = LL deflection limit,
absolute

13 = Total load deflection
limit, absolute

14 = Total camber limit,
absolute

15 = Specified camber

16 = Net area to total
area ratio

17 = Live load reduction
factor

18 = Unbraced length ratio,
y-y

19 = Unbraced length ratio,
z-z

20 = Effective length
factor, K2y

21 = Effective length
factor, K2z

22 = Moment coefficient,
kyy

23 = Moment coefficient,
kzz

24 = Bending coefficient,
Ψ

25
= Moment coefficient, kzy

26
= Moment coefficient, kyz

27 = Yield stress, Fy

28 = Compressive capacity,
Nc.Rd

29 = Tensile capacity,
Nt.Rd

30 = Bending capacity
about y-y axis,Mcy.Rd

31 = Bending capacity
about z-z axis Mcz.Rd

32 = Buckling resistance
moment, Mb.Rd

33 = Shear capacity along
z-z axis, Vz.Rd

34 = Shear capacity along
y-y axis, Vy.RD

35 = Demand/capacity ratio
limit

36 = Section class

37 = Column buckling curve,
y-y

38 = Column buckling curve,
z-z

39 = Buckling curve for LTB

40 = System overstrength factor,
Omega

41 = Is rolled section

42 = Unbraced length ratio,
LTB

43 = Effective length factor
braced, K1y

44 = Effective length factor
braced, K1z

45 = Effective length factor,
K LTB

46 = Material overstrength
factor, GammaRd

47 = Warping constant, Iw

48 = Elastic torsional buckling
force, Ncr T

49 = Elastic torsional-flexural
buckling force, Ncr TF

50
= Bending coefficient, C2

51 = Bending coefficient,
C3

52 = Warping coefficient,
kw

53 = Coordinate of load
application, za

54
= Shear center coordinate, zs

55
= Elastic critical moment for lateral-torsional buckling, Mcr

Value

The value of the considered overwrite item.

1 = Current design section

      0 or "Program
Determined"

      String corresponding
to desired shape overwrite, e.g. "W14x26" or "W14X26".

2 = Framing type

0 or "Program Default"

1 or "DCH-MRF"

2 or "DCL-MRF"

3 or "DCH-CBF"

4 or "DCL-CBF"

5 or "DCH-EBF"

6 or "DCL-EBF"

7 or "InvPendulum"

8 or "NonDissipative"

3 = Consider deflection

0 or "Program Default"

1 or "No"

2 or "Yes"

4 = Deflection check type

0 or "Program Default"

1 or "Ratio"

2 or "Absolute"

3 or "Both"

5 = DL deflection limit,
L/Value

Value >= 0; 0 means
no check for this item.

6 = SDL + LL deflection
limit, L/Value

Value >= 0; 0 means
no check for this item.

7 = LL deflection limit,
L/Value

Value >= 0; 0 means
no check for this item.

8 = Total load deflection
limit, L/Value

Value >= 0; 0 means
no check for this item.

9 = Total camber limit,
L/Value

Value >= 0; 0 means
no check for this item.

10 = DL deflection limit,
absolute

Value >= 0; 0 means
no check for this item. [L]

11 = SDL + LL deflection
limit, absolute

  Value >=
0; 0 means no check for this item. [L]

12 = LL deflection limit,
absolute

  Value >=
0; 0 means no check for this item. [L]

13 = Total load deflection
limit, absolute

  Value >=
0; 0 means no check for this item. [L]

14 = Total camber limit,
absolute

  Value >=
0; 0 means no check for this item. [L]

15 = Specified camber

  Value >=
0. [L]

16 = Net area to total
area ratio

  Value >=
0; 0 means use program default value.

17 = Live load reduction
factor

  Value >=
0; 0 means use program determined value.

18 = Unbraced length ratio,
y-y

  Value >=
0; 0 means use program determined value.

19 = Unbraced length ratio,
z-z

  Value >=
0; 0 means use program determined value.

20 = Effective length
factor sway, K2y

  Value >=
0; 0 means use program determined value.

21 = Effective length
factor sway,K2z

  Value >=
0; 0 means use program determined value.

22 = Moment coefficient,
kyy

  Value >=
0; 0 means use program determined value.

23 = Moment coefficient,
kzz

  Value >=
0; 0 means use program determined value.

24 = Bending coefficient,
Ψ

  Value >=
0; 0 means use program determined value.

25 = Moment coefficient,
kzy

  Value >=
0; 0 means use program determined value.

26 = Moment coefficient,
kyz

  Value >=
0; 0 means use program determined value.

27 = Yield stress, Fy

  Value >=
0; 0 means use program determined value. [F/L2]

28 = Compressive capacity,
Nc.Rd

  Value >=
0; 0 means use program determined value. [F]

29 = Tensile capacity,
Nt.Rd

  Value >=
0; 0 means use program determined value. [F]

30 = Bending capacity
about y-y axis, Mcy.Rd

  Value >=
0; 0 means use program determined value. [FL]

31 = Bending capacity
about z-z axis, Mcz.Rd

  Value >=
0; 0 means use program determined value. [FL]

32 = Buckling resistance
moment, Mb.Rd

  Value >=
0; 0 means use program determined value. [FL]

33 = Shear capacity along
z-z axis, Vz.Rd

  Value >=
0; 0 means use program determined value. [F]

34 = Shear capacity along
y-y axis, Vy.Rd

  Value >=
0; 0 means use program determined value. [F]

35 = Demand/capacity ratio
limit

  Value >= 0;
0 means use program determined value.

36 = Section class

0 or "Program default"

1 or "Class 1"

2 or "Class 2"

3 or "Class 3"

4 or "Class 4"

37 = Column buckling curve,
y-y

0 or "Program default"

1 or "a0"

2 or "a"

3 or "b"

4 or "c"

5 or "d"

38 = Column buckling curve,
z-z

0 or "Program default"

1 or "a0"

2 or "a"

3 or "b"

4 or "c"

5 or "d"

39 = Buckling curve for LTB

0 or "Program default"

1 or "a0"

2 or "a"

3 or "b"

4 or "c"

5 or "d"

40 = System overstrength factor,
Omega

Value >= 0; 0 means use
program determined value.

41 = Is rolled section

0 or "Program Determined"

1 or "No"

2 or "Yes"

42 = Unbraced length ratio,
LTB

Value >= 0; 0 means use
program determined value.

43 = Effective length factor
braced, K1y

Value >= 0; 0 means use
program determined value.

44 = Effective length factor
braced, K1yz

Value >= 0; 0 means use
program determined value.

45 = Effective length factor,
K LTB

Value >= 0; 0 means use
program determined value.

46 = Material overstrength
factor, GammaRd

Value >= 0; 0 means use
program determined value.

47 = Warping constant, Iw

Value >= 0; 0 means use
program determined value. [L6]

48 = Elastic torsional buckling
force, Ncr T

Value >= 0; 0 means use
program determined value. [F]

49 = Elastic torsional-flexural
buckling force, Ncr TF

Value >= 0; 0 means
use program determined value. [F]

50
= Bending coefficient, C2

Value >= 0; 0 means
use program determined value.

51 = Bending coefficient,
C3

Value >= 0; 0 means
use program determined value.

52 = Warping coefficient,
kw (used in Mcr calculation)

0.5 =<Value =<
1; 0 means use program determined value which is defaulted to 1.0.

53 = Coordinate of load
application, za (used in Mcr calculation)

54 = Shear center coordinate,
zs (used in Mcr calculation)

55
= Elastic critical moment for lateral-torsional buckling, Mcr

Value >=
0; 0 means use program determined value. [FL]

ProgDet

If this item is True, the specified value is program
determined.

## Remarks

This function retrieves the value of a steel design
preference item.

The function returns zero if the item is successfully
retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignOverwriteItemItalian_NTC_2018()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim textValue As 
 String

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

      ret = SapModel.DesignSteel.SetCode("Italian 
 NTC 2018")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start steel design

      ret = SapModel.DesignSteel.StartDesign



   'get overwrite item

      ret = SapModel.DesignSteel.Italian_NTC_2018.GetOverwrite("8", 
 2, textValue, numericValue, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v20.1.0.

Added items 50 - 53 in v22.0.0

Added items 54 - 55 in v22.0.0

Changed references from major and minor axes to y-y
and z-z in version 23.4.0.

## See Also

[SetOverwrite](SetOverwrite_%7BNTC_2018%7D.md)
