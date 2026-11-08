# SetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.BS5950\_2000.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByValItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending on the value of the ItemType item.

Item

This is an integer between 1 and 32, inclusive, indicating the overwrite item considered.

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

21 = Uniform moment factor, m Major

22 = Uniform moment factor, m Minor

23 = Eqv. uniform moment factor, mLT

24 = Yield stress, Fy

25 = Compressive capacity, Pc

26 = Tensile capacity, Pt

27 = Major bending capacity, Mc3

28 = Minor bending capacity, Mc2

29 = Buckling resistance moment, Mb

30 = Major shear capacity, Pv2

31 = Minor shear capacity, Pv3

32 = Demand/capacity ratio limit

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

21 = Uniform moment factor, m Major

Value >= 0; 0 means use program determined value.

22 = Uniform moment factor, m Minor

Value >= 0; 0 means use program determined value.

23 = Eqv. Uniform moment factor, mLT

Value >= 0; 0 means use program determined value.

24 = Yield stress, Fy

Value >= 0; 0 means use program determined value. [F/L2]

25 = Compressive capacity, Pc

Value >= 0; 0 means use program determined value. [F]

26 = Tensile capacity, Pt

Value >= 0; 0 means use program determined value. [F]

27 = Major bending capacity, Mc3

Value >= 0; 0 means use program determined value. [FL]

28 = Minor bending capacity, Mc2

Value >= 0; 0 means use program determined value. [FL]

29 = Buckling resistance moment, Mb

Value >= 0; 0 means use program determined value. [FL]

30 = Major shear capacity, Pv2

Value >= 0; 0 means use program determined value. [F]

31 = Minor shear capacity, Pv3

Value >= 0; 0 means use program determined value. [F]

32 = Demand/capacity ratio limit

Value >= 0; 0 means use program determined value.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function sets the value of a steel design overwrite item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignOverwriteItemBS5950_2000()

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("BS5950 2000")



   'set overwrite item

      ret = SapModel.DesignSteel.BS5950_2000.SetOverwrite("8", 1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetOverwrite](GetOverwrite_%7BSteel_BS5950_2000%7D.md)
