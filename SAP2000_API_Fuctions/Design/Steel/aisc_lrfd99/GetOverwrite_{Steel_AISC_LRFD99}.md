# GetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.AISC\_LRFD99.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRefProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a steel frame design procedure.

Item

This is an integer between 1 and 37, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Omega0

3 = Consider deflection

4 = Deflection check type

5 = DL deflection limit, L/Value

6 = SDL + LL deflection limit, L/Value

7 = LL deflection limit, L/Value

8 = Total load deflection limit, L/Value

9 = Total camber limit, L/Value

10 = DL deflection limit, absolute

11 = SDL + LL deflection limit, absolute

12 = LL deflection limit, absolute

13 = Total load deflection limit, absolute

14 = Total camber limit, absolute

15 = Specified camber

16 = Net area to total area ratio

17 = Live load reduction factor

18 = Unbraced length ratio, Major

19 = Unbraced length ratio, Minor Lateral Torsional Buckling

20 = Effective length factor, K Major

21 = Effective length factor, K Minor

22 = Moment coefficient, Cm Major

23 = Moment coefficient, Cm Minor

24 = Bending coefficient, Cb

25 = Non-sway moment factor, B1 Major

26 = Non-sway moment factor, B1 Minor

27 = Sway moment factor, B2 Major

28 = Sway moment factor, B2 Minor

29 = Yield stress, Fy

30 = Expected to specified Fy ratio, Ry

31 = Compressive capacity, phi\*Pnc

32 = Tensile capacity, phi\*Pnt

33 = Major bending capacity, phi\*Mn3

34 = Minor bending capacity, phi\*Mn2

35 = Major shear capacity, phi\*Vn2

36 = Minor shear capacity, phi\*Vn3

37 = Demand/capacity ratio limit

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = OMF

2 = IMF

3 = SMF

4 = OCBF

5 = SCBF

6 = EBF

2 = Omega0

Value >= 0; 0 means use a program determined value.

3 = Consider deflection

0 = Program Determined

1 = No

2 = Yes

4 = Deflection check type

0 = Program default

1 = Ratio

2 = Absolute

3 = Both

5 = DL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

6 = SDL + LL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

7 = LL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

8 = Total load deflection limit, L/Value

Value >= 0; 0 means no check for this item.

9 = Total camber limit, L/Value

Value >= 0; 0 means no check for this item.

10 = DL deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

11 = SDL + LL deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

12 = LL deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

13 = Total load deflection limit, absolute

Value >= 0; 0 means no check for this item. [L]

14 = Total camber limit, absolute

Value >= 0; 0 means no check for this item. [L]

15 = Specified camber

Value >= 0. [L]

16 = Net area to total area ratio

Value >= 0; 0 means use program default value.

17 = Live load reduction factor

Value >= 0; 0 means use program determined value.

18 = Unbraced length ratio, Major

Value >= 0; 0 means use program determined value.

19 = Unbraced length ratio, Minor Lateral Torsional Buckling

Value >= 0; 0 means use program determined value.

20 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

21 = Effective length factor, K Minor

Value >= 0; 0 means use program determined value.

22 = Moment coefficient, Cm Major

Value >= 0; 0 means use program determined value.

23 = Moment coefficient, Cm Minor

Value >= 0; 0 means use program determined value.

24 = Bending coefficient, Cb

Value >= 0; 0 means use program determined value.

25 = Non-sway moment factor, B1 Major

Value >= 0; 0 means use program determined value.

26 = Non-sway moment factor, B1 Minor

Value >= 0; 0 means use program determined value.

27 = Sway moment factor, B2 Major

Value >= 0; 0 means use program determined value.

28 = Sway moment factor, B2 Minor

Value >= 0; 0 means use program determined value.

29 = Yield stress, Fy

Value >= 0; 0 means use program determined value. [F/L2]

30 = Expected to specified Fy ratio, Ry

Value >= 0; 0 means use program determined value.

31 = Compressive capacity, phi\*Pnc

Value >= 0; 0 means use program determined value. [F]

32 = Tensile capacity, phi\*Pnt

Value >= 0; 0 means use program determined value. [F]

33 = Major bending capacity, phi\*Mn3

Value >= 0; 0 means use program determined value. [FL]

34 = Minor bending capacity, phi\*Mn2

Value >= 0; 0 means use program determined value. [FL]

35 = Major shear capacity, phi\*Vn2

Value >= 0; 0 means use program determined value. [F]

36 = Minor shear capacity, phi\*Vn3

Value >= 0; 0 means use program determined value. [F]

37 = Demand/capacity ratio limit

Value >= 0; 0 means use program determined value.

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a steel design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignOverwriteItemAISC_LRFD99()

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

      ret = SapModel.DesignSteel.SetCode("AISC-LRFD99")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start steel design

      ret = SapModel.DesignSteel.StartDesign



   'get overwrite item

      ret = SapModel.DesignSteel.AISC_LRFD99.GetOverwrite("8", 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetOverwrite](SetOverwrite_%7BSteel_AISC_LRFD99%7D.md)
