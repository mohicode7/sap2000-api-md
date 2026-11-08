# GetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.EUROCODE\_3\_1993.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRefProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a steel frame design procedure.

Item

This is an integer between 1 and 36, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Consider deflection

3 = Deflection check type

4 = DL deflection limit, L/Value

5 = SDL + LL deflection limit, L/Value

6 = LL deflection limit, L/Value

7 = Total load deflection limit, L/Value

8 = Total camber limit, L/Value

9 = DL deflection limit, absolute

10 = SDL + LL deflection limit, absolute

11 = LL deflection limit, absolute

12 = Total load deflection limit, absolute

13 = Total camber limit, absolute

14 = Specified camber

15 = Net area to total area ratio

16 = Live load reduction factor

17 = Unbraced length ratio, Major

18 = Unbraced length ratio, Minor Lateral Torsional Buckling

19 = Effective length factor, K Major

20 = Effective length factor, K Minor

21 = Moment coefficient, k Major

22 = Moment coefficient, k Minor

23 = Bending coefficient, C1

24 = Moment coefficient, k Lateral Torsional Buckling

25 = Non-sway moment factor

26 = Sway moment factor, Psi Major

27 = Sway moment factor, Psi Minor

28 = Yield stress, Fy

29 = Compressive capacity, Nc.Rd

30 = Tensile capacity, Nt.Rd

31 = Major bending capacity, Mc3.Rd

32 = Minor bending capacity, Mc2.Rd

33 = Buckling resistance moment, Mb.Rd

34 = Major shear capacity, V2.Rd

35 = Minor shear capacity, V3.RD

36 = Demand/capacity ratio limit

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = Moment Frame

2 = Braced Frame

2 = Consider deflection

0 = Program Determined

1 = No

2 = Yes

3 = Deflection check type

0 = Program default

1 = Ratio

2 = Absolute

3 = Both

4 = DL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

5 = SDL + LL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

6 = LL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

7 = Total load deflection limit, L/Value

Value >= 0; 0 means no check for this item.

8 = Total camber limit, L/Value

Value >= 0; 0 means no check for this item.

9 = DL deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

10 = SDL + LL deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

11 = LL deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

12 = Total load deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

13 = Total camber limit, absolute

Value >= 0; 0 means no check for this item. [L]

14 = Specified camber

Value >= 0. [L]

15 = Net area to total area ratio

Value >= 0; 0 means use program default value.

16 = Live load reduction factor

Value >= 0; 0 means use program determined value.

17 = Unbraced length ratio, Major

Value >= 0; 0 means use program determined value.

18 = Unbraced length ratio, Minor Lateral Torsional Buckling

Value >= 0; 0 means use program determined value.

19 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

20 = Effective length factor, K Minor

Value >= 0; 0 means use program determined value.

21 = Moment coefficient, k Major

Value >= 0; 0 means use program determined value.

22 = Moment coefficient, k Minor

Value >= 0; 0 means use program determined value.

23 = Bending coefficient, C1

Value >= 0; 0 means use program determined value.

24 = Moment coefficient, k Lateral Torsional Buckling

Value >= 0; 0 means use program determined value.

25 = Non-sway moment factor

Value >= 0; 0 means use program determined value.

26 = Sway moment factor, Psi Major

Value >= 0; 0 means use program determined value.

27 = Sway moment factor, Psi Minor

Value >= 0; 0 means use program determined value.

28 = Yield stress, Fy

Value >= 0; 0 means use program determined value. [F/L2]

29 = Compressive capacity, Nc.Rd

Value >= 0; 0 means use program determined value. [F]

30 = Tensile capacity, Nt.Rd

Value >= 0; 0 means use program determined value. [F]

31 = Major bending capacity, Mc3.Rd

Value >= 0; 0 means use program determined value. [FL]

32 = Minor bending capacity, Mc2.Rd

Value >= 0; 0 means use program determined value. [FL]

33 = Buckling resistance moment, Mb.Rd

Value >= 0; 0 means use program determined value. [FL]

34 = Major shear capacity, Vn2.Rd

Value >= 0; 0 means use program determined value. [F]

35 = Minor shear capacity, Vn3.RD

Value >= 0; 0 means use program determined value. [F]

36 = Demand/capacity ratio limit

Value >= 0; 0 means use program determined value.

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a steel design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignOverwriteItemEUROCODE_3_1993()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("EUROCODE 3-1993")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start steel design

      ret = SapModel.DesignSteel.StartDesign



   'get overwrite item

      ret = SapModel.DesignSteel.EUROCODE_3_1993.GetOverwrite("8", 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetOverwrite](SetOverwrite_%7BSteel_Eurocode_3-1993%7D.md)
