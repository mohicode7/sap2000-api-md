# GetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.ItalianNTC2008C.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item
As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a concrete frame design
procedure.

Item

This is an integer between 1 and 37, inclusive, indicating
the overwrite item considered.

1 = Framing type (beam and column)

2 = Live load reduction
factor (beam and column)

3 = Unbraced length ratio,
Major (beam and column)

4 = Unbraced length ratio,
Minor (beam and column)

5 = Effective length factor,
Beta Major (column only)

6 = Effective length factor,
Beta Minor (column only)

7 = Moment coefficient,
Cm Major (not used)

8 = Moment coefficient,
Cm Minor (not used)

9 = Nonsway moment factor,
Dns Major (not used)

10 = Nonsway moment factor,
Dns Minor (not used)

11 = Sway moment factor,
Ds Major (not used)

12 = Sway moment factor,
Ds Minor (not used)

13 = Correction factor
depending on axial load in Nominal Curvature method, Kr Major (beam and
column)

14 = Correction factor
depending on axial load in Nominal Curvature method, KPhi Major (beam
and column)

15 = Shear compressive
strut angle, TanTheta 25 = Consider torsion (beam only)

16 = Factor accounting
for creep in Nominal Curvature method, Kr Minor (beam and column)

17 = Factor accounting
for creep in Nominal Curvature method, KPhi Minor (beam and column)

18 = Coefficient depending
on the distribution of first-order moment in both Nominal Stiffness and
Nominal Curvature methods, c Major (beam and column)

19 = Coefficient depending
on the distribution of first-order moment in both Nominal Stiffness and
Nominal Curvature methods, c Minor (beam and column)

20 = Factor for contribution
of reinforcement, Ks Major (beam and column)

21 = Factor for contribution
of reinforcement, Ks Minor (beam and column)

22 = Factor for effects
of cracking, creep etc, Kc Major (beam and column)

23 = Factor for effects
of cracking, creep etc, Kc Minor (beam and column)

24 = Effective creep
coefficient, Phief (beam and column)

25 = Consider torsion
(beam only)

28 = Exposure class for crack
control

29 = Limiting crack width

30 = Age at cracking of
concrete, days

31 = Type of cement

32 = Load duration

33 = Longitudinal rebar
size top (beam only)

34 = Longitudinal rebar
size bottom (beam only)

35 = Is longitudinal rebar
ribbed?

36
= Is braced about major?

37 = Is braced about minor?

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Determined

1 = DC High

2 = DC Medium

3 = DC Low

4 = Secondary

2 = Live load reduction
factor

Value >= 0; 0 means
use program determined value.

3 = Unbraced length ratio,
Major

Value >= 0; 0 means
use program determined value.

4 = Unbraced length ratio,
Minor

Value >= 0; 0 means
use program determined value.

5 = Effective length factor,
Beta Major – only applies to column design

Value >= 0; 0 means
use program determined value.

6 = Effective length factor,
Beta Minor – only applies to column design

Value >= 0; 0 means
use program determined value.

7 = Moment coefficient,
Cm Major – not used

Value >= 0; 0 means
use program determined value.

8 = Moment coefficient,
Cm Minor – not used

Value >= 0; 0 means
use program determined value.

9 = Nonsway moment factor,
Dns Major – not used

Value >= 0; 0 means
use program determined value.

10 = Nonsway moment factor,
Dns Minor – not used

  Value >=
0; 0 means use program determined value.

11 = Sway moment factor,
Ds Major – not used

  Value >=
0; 0 means use program determined value.

12 = Sway moment factor,
Ds Minor – not used

  Value >=
0; 0 means use program determined value.

13
= Correction factor depending on axial load in Nominal Curvature method,
Kr Major

Value >= 0; 0 means
use program determined value.

14 = Correction factor
depending on axial load in Nominal Curvature method, KPhi Major

Value >= 0; 0 means
use program determined value.

15 = Shear compressive
strut angle, TanTheta

Value >= 0; 0 means
use program determined value.

16 = Factor accounting
for creep in Nominal Curvature method, Kr Minor

Value >= 0; 0 means
use program determined value.

17 = Factor accounting
for creep in Nominal Curvature method, KPhi Minor

Value >= 0; 0 means
use program determined value.

18 = Coefficient depending
on the distribution of first-order moment in both Nominal Stiffness and
Nominal Curvature methods, c Major

Value >= 0; 0 means
use program determined value.

19 = Coefficient depending
on the distribution of first-order moment in both Nominal Stiffness and
Nominal Curvature methods, c Minor

Value >= 0; 0 means
use program determined value.

20 = Factor for contribution
of reinforcement, Ks Major

Value >= 0; 0 means
use program determined value.

21 = Factor for contribution
of reinforcement, Ks Minor

Value >= 0; 0 means
use program determined value.

22 = Factor for effects
of cracking, creep etc, Kc Major

Value >= 0; 0 means
use program determined value.

23 = Factor for effects
of cracking, creep etc, Kc Minor

Value >= 0; 0 means
use program determined value.

24 = Effective creep
coefficient, Phief

Value >= 0; 0 means
use program determined value.

25 = Consider torsion

0 = No

Any other value = Yes

26
= Coefficient of concrete compressive stress limit, k1

Value >=0; 0 means use
program determined value.

27 = Coefficient of steel
tensile stress limit, k3

Value >= 0; 0 means
use program determined value.

28 = Exposure class for
crack control

0 = Program Determined

1 = X0

2 = XC1

3 = XC2

4 = XC3

5 = XC4

6 = XD1

7 = XD2

8 = XS1

9 = XS2

10 = XS3

29 = Limiting crack width

Value >= 0; 0 means
use program determined value.

30 = Age at cracking of
concrete, days

Value >= 0; 0 means
use program determined value.

31 = Type of cement

0 = Program Determined

1 = Class R

2 = Class N

3 = Class S

32 = Loading duration

0 = Program Determined

1 = Short term

2 = Long term

33
= Longitudinal rebar size top

      1
= #2

      2
= #3

      3
= #4

      4
= #5

      5
= #6

      6
= #7

      7
= #8

      8
= #9

      9
= #10

      10
= #11

      11
= #14

      12
= #18

      13
= 10M

      14
= 15M

      15
= 20M

      16
= 25M

      17
= 30M

      18
= 35M

      19
= 45M

      20
= 55M

      21
= 6d

      22
= 8d

      23
= 10d

      24
= 12d

      25
= 14d

      26
= 16d

      27
= 20d

      28
= 25d

      29
= 26d

      30
= 28d

      31
= N12

      32
= N16

      33
= N20

      34
= N24

      35
= N28

      36
= N32

      37
= N36

34 = Longitudinal rebar
size bottom

      1
= #2

      2
= #3

      3
= #4

      4
= #5

      5
= #6

      6
= #7

      7
= #8

      8
= #9

      9
= #10

      10
= #11

      11
= #14

      12
= #18

      13
= 10M

      14
= 15M

      15
= 20M

      16
= 25M

      17
= 30M

      18
= 35M

      19
= 45M

      20
= 55M

      21
= 6d

      22
= 8d

      23
= 10d

      24
= 12d

      25
= 14d

      26
= 16d

      27
= 20d

      28
= 25d

      29
= 26d

      30
= 28d

      31
= N12

      32
= N16

      33
= N20

      34
= N24

      35
= N28

      36
= N32

      37
= N36

35 = Is longitudinal rebar
ribbed?

0 = No

Any other value = Yes

36 = Is braced about major?

0
= Program Determined

1 = No

2 = Yes

37 = Is braced about minor?

0 = Program Determined

1 = No

2 = Yes

ProgDet

If this item is True, the specified value is program
determined.

## Remarks

This function retrieves the value of a concrete design
overwrite item.

The function returns zero if the item is successfully
retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignOverwriteItemItalianNTC2008C()

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



   'create new concrete frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", 
 "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288, True, "R1", "R1")



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("Italian 
 NTC 2008")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start concrete design

      ret = SapModel.DesignConcrete.StartDesign



   'get overwrite item

      ret = SapModel.DesignConcrete.ItalianNTC2008C.GetOverwrite("8", 
 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version in v23.4.0.

Added items 26 through 35 in v24.2.0

Added items 36 and 37 in v25.1.0

## See Also

[SetOverwrite](SetOverwrite%7BConcrete_NTC_2008%7D.md)
