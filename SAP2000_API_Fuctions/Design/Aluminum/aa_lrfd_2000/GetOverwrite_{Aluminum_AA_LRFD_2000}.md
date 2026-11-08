# GetOverwrite

## Syntax

SapObject.SapModel.DesignAluminum.AA\_LRFD\_2000.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a aluminum frame design procedure.

Item

This is an integer between 1 and 23, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Live load reduction factor

3 = Unbraced length ratio, Major

4 = Unbraced length ratio, Minor and Lateral Torsional Buckling

5 = Effective length factor, K Major

6 = Effective length factor, K Minor

7 = Moment coefficient, Cm Major

8 = Moment coefficient, Cm Minor

9 = Bending coefficient, Cb

10 = Buckling constant for compression, k1

11 = Buckling constant for compression, k2

12 = Buckling constant for bending, k1

13 = Buckling constant for bending, k2

14 = Safety coefficient, kt

15 = Bending coefficient, C1

16 = Bending coefficient, C2

17 = Yield stress, Fy

18 = Compressive stress, Fa

19 = Tensile stress, Ft

20 = Major bending stress, Fb3

21 = Minor bending stress, Fb2

22 = Major shear stress, Fs2

23 = Minor shear stress, Fs3

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = Moment Frame

2 = Braced Frame

2 = Live load reduction factor

Value >= 0; 0 means use a program determined value.

3 = Unbraced length ratio, Major

Value >= 0; 0 means use program determined value.

4 = Unbraced length ratio, Minor and Lateral Torsional Buckling

Value >= 0; 0 means use program determined value.

5 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

6 = Effective length factor, K Minor

Value >= 0; 0 means use program determined value.

7 = Moment coefficient, Cm Major

Value >= 0; 0 means use program determined value.

8 = Moment coefficient, Cm Minor

Value >= 0; 0 means use program determined value.

9 = Bending coefficient, Cb

Value >= 0; 0 means use program determined value.

10 = Buckling constant for compression, k1

Value >= 0; 0 means use program determined value.

11 = Buckling constant for compression, k2

Value >= 0; 0 means use program determined value.

12 = Buckling constant for bending, k1

Value >= 0; 0 means use program determined value.

13 = Buckling constant for bending, k2

Value >= 0; 0 means use program determined value.

14 = Safety coefficient, kt

Value >= 0; 0 means use program determined value.

15 = Bending coefficient, C1

Value >= 0; 0 means use program determined value.

16 = Bending coefficient, C2

Value >= 0; 0 means use program determined value.

17 = Yield stress, Fy

Value >= 0; 0 means use program determined value. [F/L2]

18 = Compressive stress, Fa

Value >= 0; 0 means use program determined value. [F/L2]

19 = Tensile stress, Ft

Value >= 0; 0 means use program determined value. [F/L2]

20 = Major bending stress, Fb3

Value >= 0; 0 means use program determined value. [F/L2]

21 = Minor bending stress, Fb2

Value >= 0; 0 means use program determined value. [F/L2]

22 = Major shear stress, Fs2

Value >= 0; 0 means use program determined value. [F/L2]

23 = Minor shear stress, Fs3

Value >= 0; 0 means use program determined value. [F/L2]

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of an aluminum design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAluminumDesignOverwriteItemAA_LRFD_2000()

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

      ret = SapModel.DesignAluminum.SetCode("AA-LRFD 2000")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start aluminum design

      ret = SapModel.DesignAluminum.StartDesign



   'get overwrite item

      ret = SapModel.DesignAluminum.AA_LRFD_2000.GetOverwrite("8", 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetOverwrite](SetOverwrite_%7BAluminum_AA_LRFD_2000%7D.md)
