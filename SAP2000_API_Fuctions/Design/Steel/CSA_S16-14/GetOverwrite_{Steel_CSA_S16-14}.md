# GetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.Canadian\_S16\_14.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item
As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a steel frame design
procedure.

Item

This is an integer between 1 and 41, inclusive, indicating
the overwrite item considered.

1 = Framing type

2 = Consider deflection

3 = Deflection check type

4 = DL deflection limit,
L/Value

5 = SDL + LL deflection
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
Major

18 = Unbraced length ratio,
Minor LTB

19 = Unbraced length ratio,
Lateral Torsional Buckling

20 = Effective length
factor, K Major

21 = Effective length
factor, K Minor

22 = Effective length
factor, K LTB

23 = Moment coefficient,
Omega1 Major

24 = Moment coefficient,
Omega1 Minor

25 = Bending coefficient,
Omega2

26 = Nonsway moment factor,
U1 Major

27 = Nonsway moment factor,
U1 Minor

28 = Sway moment factor,
U2 Major

29 = Sway moment factor,
U2 Minor

30 = Parameter for compressive
resistance, n

31 = Yield stress, Fy

32 = Expected to specified
Fy ratio, Ry

33
= Compressive resistance, Cr

34 = Tensile resistance,
Tr

35 = Major bending resistance,
Mr3

36 = Minor bending resistance,
Mr2

37 = Major shear resistance,
Vr2

38 = Minor shear resistance,
Vr3

39 = Demand/capacity ratio
limit

40 = Is HSS sections class
H?

41 = Warping torsional
constant Cw

Value

The value of the considered overwrite item.

1 = Framing type

0 = As specified in preferences

1 = Type LD MRF

2 = Type MD MRF

3 = Type D MRF

4 = Type LD CBF(V)

5 = Type LD CBF(TC)

6 = Type LD CBF(TO)

7 = Type LD CBF(OT)

8 = Type MD CBF(V)

9 = Type MD CBF(TC)

10 = Type MD CBF(TO)

11 = Type MD CBF(OT)

12 = EBF

13 = Cantilever Column

14 = Conventional MF

15 = Conventional BF

2 = Consider deflection

0 = No

Any other value = Yes

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
Major

  Value >=
0; 0 means use program determined value.

18 = Unbraced length ratio,
Minor

  Value >=
0; 0 means use program determined value.

19 = Unbraced length ratio,
Lateral Torsional Buckling

  Value >=
0; 0 means use program determined value.

20 = Effective length
factor, K Major

  Value >=
0; 0 means use program determined value.

21 = Effective length
factor, K Minor

  Value >=
0; 0 means use program determined value.

22 = Effective length
factor, K LTB

  Value >=
0; 0 means use program determined value.

23 = Moment coefficient,
Omega1 Major

  Value >=
0; 0 means use program determined value.

24 = Moment coefficient,
Omega1 Minor

  Value >=
0; 0 means use program determined value.

25 = Moment coefficient,
Omega2

  Value >=
0; 0 means use program determined value.

26 = Nonsway moment factor,
U1 Major

  Value >=
0; 0 means use program determined value.

27 = Nonsway moment factor,
U1 Minor

  Value >=
0; 0 means use program determined value.

28 = Sway moment factor,
U2 Major

  Value >=
0; 0 means use program determined value.

29 = Sway moment factor,
U2 Minor

  Value >=
0; 0 means use program determined value.

30 = Parameter for compressive
resistance, n

  Value >=
0; 0 means use program determined value.

31 = Yield stress, Fy

  Value >=
0; 0 means use program determined value. [F/L2]

32 = Expected to specified
Fy ratio, Ry

  Value >=
0; 0 means use program determined value. [F/L2]=

33 = Compressive resistance,
Cr

  Value >=
0; 0 means use program determined value. [F]

34 = Tensile resistance,
Tr

  Value >=
0; 0 means use program determined value. [F]

35 = Major bending resistance,
Mr3

  Value >=
0; 0 means use program determined value. [FL]

36 = Minor bending resistance,
Mr2

  Value >=
0; 0 means use program determined value. [FL]

37 = Major shear resistance,
Vr2

  Value >=
0; 0 means use program determined value. [F]

38 = Minor shear resistance,
Vr3

  Value >=
0; 0 means use program determined value. [F]

39 = Demand/capacity ratio
limit

  Value >=
0; 0 means use program determined value.

40 = Is HSS sections class
H?

  0 = Program
Determined

  1 = No

  2 = Yes

41 = Warping torsional
constant Cw

  Value >=
0; 0 means use program determined value.

ProgDet

If this item is True, the specified value is program
determined.

## Remarks

This function retrieves the value of a steel design
overwrite item.

The function returns zero if the item is successfully
retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignOverwriteItemCanadian_S16_14 ()

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

      ret = SapModel.DesignSteel.SetCode("CSA 
 S16-14")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start steel design

      ret = SapModel.DesignSteel.StartDesign



   'get overwrite item

      ret = SapModel.DesignSteel.Canadian_S16_14.GetOverwrite("8", 
 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 24.1.0

Added item 41 in version 25.1.0

## See Also

[SetOverwrite](SetOverwrite_%7BSteel_CSA_S16-14%7D.md)
