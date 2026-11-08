# GetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.SP63\_13330\_2012.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item
As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a concrete frame design
procedure.

Item

This is an integer between 1 and 20, inclusive, indicating
the overwrite item considered.

1 = Framing type

2 = Live load reduction
factor

3 = Unbraced length ratio,
Major

4 = Unbraced length ratio,
Minor

5 = Effective length factor,
K Major

6 = Effective length factor,
K Minor

7 = Moment amplification
factor, Eta Major

8 = Moment amplification
factor, Eta Minor

9 = Gammab3 for column

10 = Gammab3 for beam

11 = Consider torsion?

12 = (qsw,1\*Z1)/(Rs\*As,1)

13 = Corner rebar fraction
top

14 = Corner rebar fraction
bottom (beam only)

15 = Consider crack analysis?
(beam only)

16 = Crack width limit
full load (beam only)

17 = Crack width limit
long term (beam only)

18 = Longitudinal rebar
size top (beam only)

19 = Longitudinal rebar
size bottom (beam only)

20 = Is longitudinal rebar
ribbed? (beam only)

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Determined

1 = Sway

2 = Nonsway

2 = Live load reduction
factor

Value >=0; 0 means
use program determined value.

3 = Unbraced length ratio,
Major

Value >=0; 0 means
use program determined value.

4 = Unbraced length ratio,
Minor

Value >=0; 0 means
use program determined value.

5 = Effective length factor,
K Major

Value >=0; 0 means
use program determined value.

6 = Effective length factor,
K Minor

Value >=0; 0 means
use program determined value.

7 = Moment amplification
factor, Eta Major

Value >=0; 0 means
use program determined value.

8 = Moment amplification
factor, Eta Minor

Value >=0; 0 means
use program determined value.

9 = Gammab3 for column

Value >=0; 0 means
use program determined value.

10 = Gammab3 for beam

Value >=0; 0 means
use program determined value.

11 = Consider torsion?

0 = Program Determined

1 = No

2 = Yes

12 = (qsw,1\*Z1)/(Rs\*As,1)

0.5 <= Value <=
1.5; 0 means use program determined value.

13 = Corner rebar fraction
top

0 < Value <= 1;
0 means use program determined value.

14 = Corner rebar fraction
bottom

0 < Value <= 1;
0 means use program determined value.

14 = Corner rebar fraction
bottom

0 < Value <= 1;
0 means use program determined value.

15 = Consider crack analysis?

0 = No

Any other value = Yes

16 = Crack width limit
full load

Value >=0; 0 means use
program determined value.

17 = Crack width limit
long term

Value >=0; 0 means use
program determined value.

18 = Longitudinal rebar
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

19 = Longitudinal rebar
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

20 = Is longitudinal rebar
ribbed?

0 = No

Any other value = Yes

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
Sub GetConcreteDesignOverwriteItemSP63_13330_2012()

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

      ret = SapModel.DesignConcrete.SetCode("SP 
 63.13330.2012")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start concrete design

      ret = SapModel.DesignConcrete.StartDesign



   'get overwrite item

      ret = SapModel.DesignConcrete.SP63_13330_2012.GetOverwrite("8", 
 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v20.1.0.

Added items 15-20 in v23.4.0

## See Also

[SetOverwrite](SetOverwrite.md)
