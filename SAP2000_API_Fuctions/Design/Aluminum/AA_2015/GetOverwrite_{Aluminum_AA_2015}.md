# GetOverwrite

## Syntax

SapObject.SapModel.DesignAluminum.AA\_2015.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with aluminum frame design procedure.

Item

This is an integer between 1 and 27, inclusive, indicating the overwrite item considered.

1 = Live load reduction factor

2 = Unbraced length ratio, Major

3 = Unbraced length ratio, Minor

4 = Unbraced length, Lateral Torsional Buckling (LTB)

5 = Effective length factor, K Major

6 = Effective length factor, K Minor

7 = Effective length factor, K LTB

8 = Bending coefficient, Cb

9 = Buckling constant for compression, k1

10 = Buckling constant for compression, k2

11 = Buckling constant for bending, k1

12 = Buckling constant for bending, k2

13 = Safety coefficient, kt

14 = Bending coefficient, C1

15 = Bending coefficient, C2

16 = Net area over gross area ratio

17 = Buckling constant, Ct

18 = Coordinate of load application, Za

19 = Demand/Capacity ratio limit

20 = Effective length factor, K1 Major

21 = Effective length factor, K1 Minor

22 = Moment coefficient factor, Cm Major

23 = Moment coefficient factor, Cm Minor

24 = Nonsway moment factor, B1 Major

25 = Nonsway moment factor, B1 Minor

26 = Sway moment factor, B2 Major

27 = Sway moment factor, B2 Minor

Value

The value of the considered overwrite item.

1 = Live load reduction factor

Value >= 0; 0 means use a program determined value.

2 = Unbraced length ratio, Major

Value >= 0; 0 means use program determined value.

3 = Unbraced length ratio, Minor

Value >= 0; 0 means use program determined value.

4 = Unbraced length ratio, Lateral Torsional Buckling

Value >= 0; 0 means use program determined value.

5 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

6 = Effective length factor, K Minor

Value >= 0; 0 means use program determined value.

7 = Effective length factor, K LTB

Value >= 0; 0 means use program determined value.

8 = Bending coefficient, Cb

Value >= 0; 0 means use program determined value.

9 = Buckling constant for compression, k1

Value >= 0; 0 means use program determined value.

10 = Buckling constant for compression, k2

Value >= 0; 0 means use program determined value.

11 = Buckling constant for bending, k1

Value >= 0; 0 means use program determined value.

12 = Buckling constant for bending, k2

Value >= 0; 0 means use program determined value.

13 = Safety coefficient, kt

Value >= 0; 0 means use program determined value.

14 = Bending coefficient, C1

Value >= 0; 0 means use program determined value.

15 = Bending coefficient, C2

Value >= 0; 0 means use program determined value.

16 = Net area over gross area ratio

Value >= 0; 0 means use program determined value.

17 = Buckling constant, Ct

Value >= 0; 0 means use program determined value.

18 = Coordinate of load application, Za

Value >= 0; 0 means use program determined value.

19 = Demand/capacity ratio limit

Value >= 0; 0 means use program determined value.

20 = Effective length factor, K1 Major

Value >= 0; 0 means use program determined value.

21 = Effective length factor, K1 Minor

Value >= 0; 0 means use program determined value.

22 = Moment coefficient, Cm Major

Value >= 0; 0 means use program determined value.

23 = Moment coefficient, Cm Minor

Value >= 0; 0 means use program determined value.

24 = Nonsway moment factor, B1 Major

Value >= 0; 0 means use program determined value.

25 = Nonsway moment factor, B1 Minor

Value >= 0; 0 means use program determined value.

26 = Sway moment factor, B2 Major

Value >= 0; 0 means use program determined value.

27 = Sway moment factor, B2 Minor

Value >= 0; 0 means use program determined value.

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of an aluminum design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAluminumDesignOverwriteItemAA_2015()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim Name As String

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



       'add aluminum material

          ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_ALUMINUM, , , MATERIAL_ALUMINUM_SUBTYPE_6061_T6)



       'create new aluminum frame section property

          ret = SapModel.PropFrame.SetISection("AI", Name , 18, 6, 0.5, 0.3, 6, 0.5)



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "AI", "AI")



       'set aluminum design code

          ret = SapModel.DesignAluminum.SetCode("AA 2015")



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'start aluminum design

          ret = SapModel.DesignAluminum.StartDesign



       'get overwrite item

          ret = SapModel.DesignAluminum.AA_2015.GetOverwrite("8", 1, Value, ProgDet)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 22.1.0

Added items 20 through 27 in version 27.0.0

## See Also

[SetOverwrite](SetOverwrite_%7BAluminum_AA_2015%7D.md)
