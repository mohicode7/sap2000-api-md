# SetPreference

## Syntax

SapObject.SapModel.DesignConcrete.SP63\_13330\_2012.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value
As Double) As Long

## Parameters

Item

This is an integer between 1 and 35, inclusive, indicating
the preference item considered.

1 = Multi-response case
design

2 = Number of interaction
curves

3 = Number of interaction
points

4 = Consider minimum eccentricity?

5 = Consider torsion?

6 = (qsw,1\*Z1)/(Rs\*As,1)

7 = Corner rebar fraction
top

8 = Corner rebar fraction
bottom

9 = Relative humidity

10 = Moisture content

11 = Gamma\_b

12 = Gamma\_bt

13 = Gamma\_b1 short term

14 = Gamma\_b1 long term

15 = Gamma\_b2

16 = Gamma\_b3 beams

17 = Gamma\_b3 columns

18 = Gamma\_b4

19 = Gamma\_b5

20 = Gamma\_S

21 = Gamma\_S1

22 = Pattern live load
factor

23 = Utilization factor
limit

24 = Live load duration
factor

25 = Snow load duration
factor

26 = Consider crack analysis

27 = Crack width limit
full load

28 = Crack width limit
long term

29 = Longitudinal rebar
size top

30 = Longitudinal rebar
size bottom

31 = Is longitudinal rebar
ribbed?

32 = Reliability factor

33 = Seismic factor mtr\_strength

34 = Seismic factor mtr\_shear

35 = Site seismicity

Value

The value of the considered preference item.

1 = Multi-response case
design

      1
= Envelopes

      2
= Step-by-step

      3
= Last step

      4
= Envelopes -- All

      5
= Step-by-step -- All

2 = Number of interaction
curves

      Value
>= 4 and divisible by 4

3 = Number of interaction
points

Value >= 5 and odd

4 = Consider minimum eccentricity?

0 = No

Any other value = Yes

5 = Consider torsion?

0 = No

Any other value = Yes

6 = (qsw,1\*Z1)/(Rs\*As,1)

0.5 <= Value <=
1.5

7 = Corner rebar fraction
top

0 < Value <= 1

8 = Corner rebar fraction
bottom

0 < Value <= 1

9 = Relative humidity

0 <= Value <= 100

10 = Moisture content

0 <= Value <= 1

11 = Gamma\_b

Value > 0

12 = Gamma\_bt

Value > 0

13 = Gamma\_b1 short term

Value > 0

14 = Gamma\_b1 long term

Value > 0

15 = Gamma\_b2

Value > 0

16 = Gamma\_b3 beams

Value > 0

17 = Gamma\_b3 columns

Value > 0

18 = Gamma\_b4

Value > 0

19 = Gamma\_b5

Value > 0

20 = Gamma\_S

Value > 0

21 = Gamma\_S1

Value > 0

22 = Pattern live load
factor

Value >= 0

23 = Utilization factor
limit

Value > 0

24 = Live load duration
factor

Value >= 0

25 = Snow load duration
factor

Value >= 0

26 = Consider crack analysis?

0 = No

Any other value = Yes

27 = Crack width limit
full load

Value >= 0

28 = Crack width limit
long term

Value >= 0

29 = Longitudinal rebar
size top

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

30
= Longitudinal rebar size bottom

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

31 = Is longitudinal rebar
ribbed?

0 = No

Any other value = Yes

32 = Reliability factor

Value > 0

33 = Seismic factor mtr\_strength

Value > 0

34 = Seismic factor mtr\_shear

Value > 0

35 = Site seismicity

      1
= Site seismicity 9

      2
= Site seismicity 8

      3
= Site seismicity 7

      4
= Non seismic

## Remarks

This function retrieves the value of a concrete design
preference item.

The function returns zero if the item is successfully
retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignPreferenceItemSP63_13330_2012()

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



   'create new concrete frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", 
 "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288, True,"R1", "R1")



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("SP63.13330.2012")



   'set preference item

      ret = SapModel.DesignConcrete.SP63_13330_2012.SetPreference(11, 
 1.2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v20.1.0.

Added items 24-35 in v23.4.0

## See Also

[GetPreference](GetPreference.md)
