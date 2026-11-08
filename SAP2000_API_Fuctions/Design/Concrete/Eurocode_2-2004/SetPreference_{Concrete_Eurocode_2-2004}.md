# SetPreference

## Syntax

SapObject.SapModel.DesignConcrete.Eurocode\_2\_2004.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value
As Double) As Long

## Parameters

Item

This is an integer between 1 and 23, inclusive, indicating
the preference item considered.

1 = Country

2 = Combos equation

3 = Second order method

4 = Number of interaction
curves

5 = Number of interaction
points

6 = Consider minimum eccentricity

7 = Theta0

8 = Gamma steel

9 = Gamma concrete

10 = AlphaCC

11 = AlphaCT

12 = AlphaLCC

13 = AlphaLCT

14 = Pattern live load
factor

15 = Utilization factor
limit

16 = Multi-response case design

17 = Reliability Class

18 = GammacE

19 = Alphae
(not used)

20 = Consider Torsion

21 = Longitudinal rebar
size top

22 = Longitudinal rebar
size bottom

23 = Is longitudinal rebar
ribbed?

Value

The value of the considered preference item.

1 = Country

      1
= CEN Default

      2
= United Kingdom

      3
= Slovenia

      5
= Norway

      6
= Singapore

      7
= Sweden

      8
= Finland

      9
= Denmark

    10
= Portugal

    11
= Germany

    12
= Poland

    13
= Ireland

2 = Combos equation

      1
= Eq. 6.10

      2
= Max of Eqs. 6.10a and 6.10b

3 = Second order method

      1
= Nominal stiffness

      2
= Nominal curvature

      3
= None

4 = Number of interaction
curves

Value >= 4 and divisible
by 4

5 = Number of interaction
points

Value >= 5 and odd

6 = Consider minimum eccentricity

0 = No

Any other value = Yes

7 = Theta0

Value > 0

8 = Gamma steel

Value > 0

9 = Gamma concrete

Value > 0

10
= AlphaCC

  Value
> 0

11
= AlphaCT

  Value
> 0

12
= AlphaLCC

  Value
> 0

13
= AlphaLCT

  Value
> 0

14 = Pattern live load
factor

  Value >=
0

15 = Utilization factor
limit

  Value >
0

16 = Multi-response case
design

  1 = Envelopes

  2 = Step-by-step

  3 = Last
step

  4 = Envelopes
-- All

  5 = Step-by-step
-- All

17
= Reliability Class

 1 = Class 1

 2 = Class 2

 3 = Class 3

                       18
= GammacE

                             Value
> 0

19
= Alphae (not used)

                             Value
> 0

20 = Consider torsion

0 = No

Any
other value = Yes

21 = Longitudinal rebar
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

22 = Longitudinal rebar
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

23 = Is longitudinal rebar
ribbed?

0 = No

Any other value = Yes

## Remarks

This function sets the value of a concrete design preference
item.

The function returns zero if the item is successfully
set; otherwise, it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignPreferenceItemEurocode_2_2004()

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

      ret = SapModel.DesignConcrete.SetCode("Eurocode 
 2-2004")



   'set preference item

      ret = SapModel.DesignConcrete.Eurocode_2_2004.SetPreference(5, 
 9)



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

Added items 18 – 19 in v21.1.0

Added item 20 in v23.4.0.

Changed item 19
to Not Used in v24.2.0

Added items 21 – 23 in v24.2.0

## See Also

[GetPreference](GetPreference%7BConcrete_Eurocode_2-2004%7D.md)
