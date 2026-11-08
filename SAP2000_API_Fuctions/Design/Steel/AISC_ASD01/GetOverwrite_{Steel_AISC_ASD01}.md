# GetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.AISC\_ASD01.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRefProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a steel frame design procedure.

Item

This is an integer between 1 and 32, inclusive, indicating the overwrite item considered.

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

25 = Yield stress, Fy

26 = Compressive stress, Fa

27 = Tensile stress, Ft

28 = Major bending stress, Fb3

29 = Minor bending stress, Fb2

30 = Major shear stress, Fv2

31 = Minor shear stress, Fv3

32 = Demand/capacity ratio limit

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

25 = Yield stress, Fy

Value >= 0; 0 means use program determined value. [F/L2]

26 = Compressive stress, Fa

Value >= 0; 0 means use program determined value. [F/L2]

27 = Tensile stress, Ft

Value >= 0; 0 means use program determined value. [F/L2]

28 = Major bending stress, Fb3

Value >= 0; 0 means use program determined value. [F/L2]

29 = Minor bending stress, Fb2

Value >= 0; 0 means use program determined value. [F/L2]

30 = Major shear stress, Fv2

Value >= 0; 0 means use program determined value. [F/L2]

31 = Minor shear stress, Fv3

Value >= 0; 0 means use program determined value. [F/L2]

32 = Demand/capacity ratio limit

Value >= 0; 0 means use program determined value.

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a steel design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignOverwriteItemAISC_ASD01()

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

      ret = SapModel.DesignSteel.SetCode("AISC-ASD01")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start steel design

      ret = SapModel.DesignSteel.StartDesign



   'get overwrite item

      ret = SapModel.DesignSteel.AISC_ASD01.GetOverwrite("8", 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetOverwrite](SetOverwrite_%7BSteel_AISC_ASD01%7D.md)
