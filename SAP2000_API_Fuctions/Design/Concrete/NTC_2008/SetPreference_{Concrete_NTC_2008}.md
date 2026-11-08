# SetPreference

## Syntax

SapObject.SapModel.DesignConcrete.ItalianNTC2008C.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value
As Double) As Long

## Parameters

Item

This is an integer between 3 and 23, inclusive, indicating
the preference item considered.

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

17 = GammacE

18 = Alphae (not used)

20 = Consider torsion

21
= Longitudinal rebar size top

22 = Longitudinal rebar
size bottom

23 = Is longitudinal rebar
ribbed?

Value

The value of the considered preference item.

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
= GammacE

                             Value
> 0

18 = Alphae

                             Value
> 0

20
= Consider torsion

0 = No

Any other value = Yes

21
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
Sub SetConcreteDesignPreferenceItemItalianNTC2008C()

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

      ret = SapModel.DesignConcrete.SetCode("Italian 
 NTC 2008")



   'set preference item

      ret = SapModel.DesignConcrete.ItalianNTC2008C.SetPreference(5, 
 9)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version in v23.4.0.

Added items 21 – 23 in v24.2.0

## See Also

[GetPreference](GetPreference_%7BConcrete_NTC_2008%7D.md)
