# GetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.Chinese\_2002.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRefProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a steel frame design procedure.

Item

This is an integer between 1 and 51, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Element type

3 = Is transfer column

4 = Seismic magnification factor

5 = Is rolled section

6 = Is flange edge cut by gas

7 = Is both end pinned

8 = Ignore b/t check

9 = Classify beam as flexo-compression member

10 = Is beam top loaded

11 = Consider deflection

12 = Deflection check type

13 = DL deflection limit, L/Value

14 = SDL + LL deflection limit, L/Value

15 = LL deflection limit, L/Value

16 = Total load deflection limit, L/Value

17 = Total camber limit, L/Value

18 = DL deflection limit, absolute

19 = SDL + LL deflection limit, absolute

20 = LL deflection limit, absolute

21 = Total load deflection limit, absolute

22 = Total camber limit, absolute

23 = Specified camber

24 = Net area to total area ratio

25 = Live load reduction factor

26 = Unbraced length ratio, Major

27 = Unbraced length ratio, Minor Lateral Torsional Buckling

28 = Effective length factor, Mue Major

29 = Effective length factor, Mue Minor

30 = Moment coefficient, Beta\_m Major

31 = Moment coefficient, Beta\_m Minor

32 = Moment coefficient, Beta\_t Major

33 = Moment coefficient, Beta\_t Minor

34 = Axial stability coefficient, Phi Major

35 = Axial stability coefficient, Phi Minor

36 = Flexural stability coeff, Phi\_bMajor

37 = Flexural stability coeff, Phi\_bMinor

38 = Plasticity factor, Gamma Major

39 = Plasticity factor, Gamma Minor

40 = Section influence coefficient, Eta

41 = B/C capacity factor, Eta

42 = Euler moment factor, Delta Major

43 = Euler moment factor, Delta Minor

44 = Yield stress, Fy

45 = Allowable normal stress, f

46 = Allowable shear stress, fv

47 = Consider fictitious shear

48 = Demand/capacity ratio limit

49 = Dual system magnification factor

50 = Lo/r limit in compression

51 = L/r limit in tension

Value

The value of the considered overwrite item.

1 = Framing type

0 = As specified in preferences

1 = Sway Moment Frame, SMF

2 = Concentrically Braced Frame, CBF

3 = Eccentrically Braced Frame, EBF

4 = NonSway Moment Frame, NMF

2 = Element type

0 = Program Determined

1 = Column

2 = Beam

3 = Brace

4 = Truss

3 = Is transfer column

0 = Program Determined

1 = No

2 = Yes

4 = Seismic magnification factor

Value >= 0; 0 means no check for this item.

5 = Is rolled section

0 = Program Determined

1 = No

2 = Yes

6 = Is flange edge cut by gas

0 = Program Determined

1 = No

2 = Yes

7 = Is both end pinned

0 = Program Determined

1 = No

2 = Yes

8 = Ignore b/t check

0 = Program Determined

1 = No

2 = Yes

9 = Classify beam as flexo-compression member

0 = Program Determined

1 = No

2 = Yes

10 = Is beam top loaded

0 = Program Determined

1 = No

2 = Yes

11 = Consider deflection

0 = Program Determined

1 = No

2 = Yes

12 = Deflection check type

0 = Program default

1 = Ratio

2 = Absolute

3 = Both

13 = DL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

14 = SDL + LL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

15 = LL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

16 = Total load deflection limit, L/Value

Value >= 0; 0 means no check for this item.

17 = Total camber limit, L/Value

Value >= 0; 0 means no check for this item.

18 = DL deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

19 = SDL + LL deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

20 = LL deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

21 = Total load deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

22 = Total camber limit, absolute

Value >= 0; 0 means no check for this item. [L]

23 = Specified camber

Value >= 0. [L]

24 = Net area to total area ratio

Value >= 0; 0 means use program default value.

25 = Live load reduction factor

Value >= 0; 0 means use program determined value.

26 = Unbraced length ratio, Major

Value >= 0; 0 means use program determined value.

27 = Unbraced length ratio, Minor Lateral Torsional Buckling

Value >= 0; 0 means use program determined value.

28 = Effective length factor, Mue Major

Value >= 0; 0 means use program determined value.

29 = Effective length factor, Mue Minor

Value >= 0; 0 means use program determined value.

30 = Moment coefficient, Beta\_m Major

Value >= 0; 0 means use program determined value.

31 = Moment coefficient, Beta\_m Minor

Value >= 0; 0 means use program determined value.

32 = Moment coefficient, Beta\_t Major

Value >= 0; 0 means use program determined value.

33 = Moment coefficient, Beta\_t Minor

Value >= 0; 0 means use program determined value.

34 = Axial stability coefficient, Phi Major

Value >= 0; 0 means use program determined value.

35 = Axial stability coefficient, Phi Minor

Value >= 0; 0 means use program determined value.

36 = Flexural stability coefficient, Phi\_b Major

Value >= 0; 0 means use program determined value.

37 = Flexural stability coefficient, Phi\_b Minor

Value >= 0; 0 means use program determined value.

38 = Plasticity factor, Gamma Major

Value >= 0; 0 means use program determined value.

39 = Plasticity factor, Gamma Minor

Value >= 0; 0 means use program determined value.

40 = Section influence coefficient, Eta

Value >= 0; 0 means use program determined value.

41 = B/C capacity factor, Eta

Value >= 0; 0 means use program determined value.

42 = Euler moment factor, Delta Major

Value >= 0; 0 means use program determined value.

43 = Euler moment factor, Delta Minor

Value >= 0; 0 means use program determined value.

44 = Yield stress, Fy

Value >= 0; 0 means use program determined value. [F/L2]

45 = Allowable normal stress, f

Value >= 0; 0 means use program determined value. [F/L2]

46 = Allowable shear stress, fv

Value >= 0; 0 means use program determined value. [F/L2]

47 = Consider fictitious shear

0 = Program Determined

1 = No

2 = Yes

48 = Demand/capacity ratio limit

Value >= 0; 0 means use program determined value.

49 = Dual system magnification factor

        Value >= 0; 0 means use program default value.

50 = Lo/r limit in compression

        Value >= 0; 0 means use program determined value.

51 = L/r limit in tension

        Value >= 0; 0 means use program determined value.

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a steel design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignOverwriteItemChinese_2002()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value As Double

      Dim ProgDet As Boolean



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("Chinese 2002")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start steel design

      ret = SapModel.DesignSteel.StartDesign



   'get overwrite item

      ret = SapModel.DesignSteel.Chinese_2002.GetOverwrite("8", 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Added Items 49, 50, and 51 in Version 14.0.0.

Modified Item 1 and added Truss to Item 2 in version 14.1.0.

## See Also

[SetOverwrite](SetOverwrite_%7BSteel_Chinese_2002%7D.md)
