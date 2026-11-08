# GetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.TCVN\_5574\_2018.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a concrete frame design procedure.

Item

This is an integer between 1 and 21, inclusive, indicating the overwrite item considered.

1 = Framing type (beam and column)

2 = Live load reduction factor (beam and column)

3 = Unbraced length ratio, Major (beam and column)

4 = Unbraced length ratio, Minor (beam and column)

5 = Effective length factor, K Major (columns only)

6 = Effective length factor, K Minor (columns only)

7 = Moment amplification factor, Eta Major (not used)

8 = Moment amplification factor, Eta Minor (not used)

9 = Gammab3 for column (columns only)

10 = Gammab3 for beam (columns only)

11 = Consider torsion? (beams only)

12 = (qsw,1\*Z1) / (Rs\*As,1)

13 = Corner rebar fraction top (beams only)

14 = Corner rebar fraction bottom (beams only)

15 = Consider minimum eccentricity (beams only)

16 = Consider crack analysis? (beams only)

17 = Crack width limit full load (beams only)

18 = Crack width limit long term (beams only)

19 = Longitudinal rebar size top (beams only)

20 = Longitudinal rebar size bottom (beams only)

21 = Is longitudinal rebar ribbed? (beams only)

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Determined

1 = DC High

2 = DC Medium

3 - DC Low

4 Secondary

2 = Live load reduction factor

Value >=0; 0 means use program determined value.

3 = Unbraced length ratio, Major

Value >=0; 0 means use program determined value.

4 = Unbraced length ratio, Minor

Value >=0; 0 means use program determined value.

5 = Effective length factor, K Major - only applies to column design

Value >=0; 0 means use program determined value.

6 = Effective length factor, K Minor - only applies to column design

Value >=0; 0 means use program determined value.

7 = Moment amplification factor, Eta Major

Value >=0; 0 means use program determined value.

8 = Moment amplification factor, Eta Minor

Value >=0; 0 means use program determined value.

9 = Gammab3 (columns only)

Value >=0; 0 means use program determined value.

10 = Gammab3 (beams only)

Value >=0; 0 means use program determined value.

11 = Consider torsion? (beams only)

0 = Program Determined

Any other value = Yes

12 = (qsw,1\*Z1) / (Rs\*As,1)

0.5 <= Value <= 1.5; 0 means use program determined value.

13 = Corner rebar fraction top  (beams only)

0 < Value <= 1; 0 means use program determined value.

14 = Corner rebar fraction bottom  (beams only)

0 < Value <= 1; 0 means use program determined value.

15 = Consider minimum eccentricity  (columns only)

0 = No

Any other value = Yes

16= Consider crack analysis?

0 = No

Any other value = Yes

17 = Crack width limit full load (beams only)

Value >=0; 0 means use program determined value.

18 = Crack width limit long term (beams only)

Value >=0; 0 means use program determined value.

19 = Longitudinal rebar size top

      1 = #2

      2 = #3

      3 = #4

      4 = #5

      5 = #6

      6 = #7

      7 = #8

      8 = #9

      9 = #10

      10 = #11

      11 = #14

      12 = #18

      13 = 10M

      14 = 15M

      15 = 20M

      16 = 25M

      17 = 30M

      18 = 35M

      19 = 45M

      20 = 55M

      21 = 6d

      22 = 8d

      23 = 10d

      24 = 12d

      25 = 14d

      26 = 16d

      27 = 20d

      28 = 25d

      29 = 26d

      30 = 28d

      31 = N12

      32 = N16

      33 = N20

      34 = N24

      35 = N28

      36 = N32

      37 = N36

20 = Longitudinal rebar size bottom

      1 = #2

      2 = #3

      3 = #4

      4 = #5

      5 = #6

      6 = #7

      7 = #8

      8 = #9

      9 = #10

      10 = #11

      11 = #14

      12 = #18

      13 = 10M

      14 = 15M

      15 = 20M

      16 = 25M

      17 = 30M

      18 = 35M

      19 = 45M

      20 = 55M

      21 = 6d

      22 = 8d

      23 = 10d

      24 = 12d

      25 = 14d

      26 = 16d

      27 = 20d

      28 = 25d

      29 = 26d

      30 = 28d

      31 = N12

      32 = N16

      33 = N20

      34 = N24

      35 = N28

      36 = N32

      37 = N36    

21 = Is longitudinal rebar ribbed?

0 = No

Any other value = Yes

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a concrete design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignOverwriteItemTCVN55742018()

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

          ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



       'set concrete design code

          ret = SapModel.DesignConcrete.SetCode("TCVN 5574:2018")



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'start concrete design

          ret = SapModel.DesignConcrete.StartDesign



       'get overwrite item

          ret = SapModel.DesignConcrete.TCVN_5574_2018.GetOverwrite("8", 1, Value, ProgDet)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in v27.0.0.

## See Also

[SetOverwrite](SetOverwrite.md)
