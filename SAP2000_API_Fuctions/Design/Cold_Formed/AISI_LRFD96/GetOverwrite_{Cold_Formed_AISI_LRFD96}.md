# GetOverwrite

## Syntax

SapObject.SapModel.DesignColdFormed.AISI\_LRFD96.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a cold formed frame design procedure.

Item

This is an integer between 1 and 27, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Live load reduction factor

3 = Unbraced length ratio, Major

4 = Unbraced length ratio, Minor and Lateral Torsional Buckling

5 = Effective length factor, K Major

6 = Effective length factor, K Minor

7 = Moment coefficient, Cm Major

8 = Moment coefficient, Cm Minor

9 = Moment coefficient, Ctf Major

10 = Moment coefficient, Ctf Minor

11 = Bending coefficient, Cb

12 = Moment factor, Alpha Major

13 = Moment factor, Alpha Minor

14 = Through fastened to deck

15 = Fastener eccentricity, a/b

16 = Hole diameter at top flange

17 = Hole diameter at bottom flange

18 = Hole diameter on web

19 = Yield stress, Fy

20 = Nominal compressive capacity, Pnc

21 = Nominal tensile capacity, Pnt

22 = Nominal bending capacity for yielding, Mn33

23 = Nominal bending capacity for yielding, Mn22

24 = Nominal bending capacity for lateral torsional buckling, Mn33

25 = Nominal bending capacity for lateral torsional buckling, Mn22

26 = Nominal shear capacity, Vn2

27 = Nominal shear capacity, Vn3

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

9 = Moment coefficient, Ctf Major

Value >= 0; 0 means use program determined value.

10 = Moment coefficient, Ctf Minor

Value >= 0; 0 means use program determined value.

11 = Bending coefficient, Cb

Value >= 0; 0 means use program determined value.

12 = Moment factor, Alpha Major

Value >= 0; 0 means use program determined value.

13 = Moment factor, Alpha Minor

Value >= 0; 0 means use program determined value.

14 = Through fastened to deck

0 = No.

Any other value = Yes.

15 = Fastener eccentricity, a/b

Value >= 0; 0 means use program determined value. [L]

16 = Hole diameter at top flange

Value >= 0; 0 means use program determined value. [L]

17 = Hole diameter at bottom flange

Value >= 0; 0 means use program determined value. [L]

18 = Hole diameter on web

Value >= 0; 0 means use program determined value. [L]

19 = Yield stress, Fy

Value >= 0; 0 means use program determined value. [F/L2]

20 = Nominal compressive capacity, Pnc

Value >= 0; 0 means use program determined value. [F]

21 = Nominal tensile capacity, Pnt

Value >= 0; 0 means use program determined value. [F]

22 = Nominal bending capacity for yielding, Mn33

Value >= 0; 0 means use program determined value. [FL]

23 = Nominal bending capacity for yielding, Mn22

Value >= 0; 0 means use program determined value. [FL]

24 = Nominal bending capacity for lateral torsional buckling, Mn33

Value >= 0; 0 means use program determined value. [FL]

25 = Nominal bending capacity for lateral torsional buckling, Mn22

Value >= 0; 0 means use program determined value. [FL]

26 = Nominal shear capacity, Vn2

Value >= 0; 0 means use program determined value. [F]

27 = Nominal shear capacity, Vn3

Value >= 0; 0 means use program determined value. [F]

ProgDet

If this item is True, the specified value is program determined.

## Remarks

This function retrieves the value of a cold formed design overwrite item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetColdFormedDesignOverwriteItemAISI_LRFD96()

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



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section property

      ret = SapModel.PropFrame.SetColdC("CdC", Name , 9, 3, 0.06, 0.25, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 72, 2, 144, True, "CdC", "CdC")



   'set cold formed design code

      ret = SapModel.DesignColdFormed.SetCode("AISI-LRFD96")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start cold formed design

      ret = SapModel.DesignColdFormed.StartDesign



   'get overwrite item

      ret = SapModel.DesignColdFormed.AISI_LRFD96.GetOverwrite("8", 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetOverwrite](SetOverwrite_%7BCold_Formed_AISI_LRFD96%7D.md)
