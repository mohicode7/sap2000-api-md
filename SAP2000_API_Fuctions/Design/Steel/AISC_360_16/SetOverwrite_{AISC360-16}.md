# SetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.AISC360\_16.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item
As Long, ByVal Value As Double, Optional ByVal ItemType As eItemType =
Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending
on the value of the ItemType item.

Item

This is an integer between 1 and 43, inclusive, indicating
the overwrite item considered.

1 = Framing type

2 = Omega0

3 = Consider deflection

4 = Deflection check type

5 = DL deflection limit,
L/Value

6 = SDL + LL deflection
limit, L/Value

7 = LL deflection limit,
L/Value

8 = Total load deflection
limit, L/Value

9 = Total camber limit,
L/Value

10 = DL deflection limit,
absolute

11 = SDL + LL deflection
limit, absolute

12 = LL deflection limit,
absolute

13 = Total load deflection
limit, absolute

14 = Total camber limit,
absolute

15 = Specified camber

16 = Net area to total
area ratio

17 = Live load reduction
factor

18 = Unbraced length ratio,
Major

19 = Unbraced length ratio,
Minor

20 = Unbraced length ratio,
Lateral Torsional Buckling

21 = Effective length
factor, K1 Major

22 = Effective length
factor, K1 Minor

23 = Effective length
factor, K2 Major

24 = Effective length
factor, K2 Minor

25 = Effective length
factor, K Lateral Torsional Buckling

26 = Moment coefficient,
Cm Major

27 = Moment coefficient,
Cm Minor

28 = Bending coefficient,
Cb

29 = Nonsway moment factor,
B1 Major

30 = Nonsway moment factor,
B1 Minor

31 = Sway moment factor,
B2 Major

32 = Sway moment factor,
B2 Minor

33 = Reduce HSS thickness

34 = HSS welding type

35 = Yield stress, Fy

36 = Expected to specified
Fy ratio, Ry

37 = Compressive capacity,
Pnc

38 = Tensile capacity,
Pnt

39 = Major bending capacity,
Mn3

40 = Minor bending capacity,
Mn2

41 = Major shear capacity,
Vn2

42 = Minor shear capacity,
Vn3

43 = Demand/capacity ratio
limit

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = SMF

2 = IMF

3 = OMF

4 = SCBF

5 = OCBF

6 = OCBFI

7 = EBF

2 = Omega0

Value >= 0; 0 means
use a program determined value.

3 = Consider deflection

0 = Program default

1 = No

2 = Yes

4 = Deflection check type

0 = Program default

1 = Ratio

2 = Absolute

3 = Both

5 = DL deflection limit,
L/Value

Value >= 0; 0 means
no check for this item.

6 = SDL + LL deflection
limit, L/Value

Value >= 0; 0 means
no check for this item.

7 = LL deflection limit,
L/Value

Value >= 0; 0 means
no check for this item.

8 = Total load deflection
limit, L/Value

Value >= 0; 0 means
no check for this item.

9 = Total camber limit,
L/Value

Value >= 0; 0 means
no check for this item.

10 = DL deflection limit,
absolute

Value >= 0; 0 means
no check for this item. [L]

11 = SDL + LL deflection
limit, absolute

Value >= 0; 0 means
no check for this item. [L]

12 = LL deflection limit,
absolute

Value >= 0; 0 means
no check for this item. [L]

13 = Total load deflection
limit, absolute

Value >= 0; 0 means
no check for this item. [L]}

14 = Total camber limit,
absolute

Value >= 0; 0 means
no check for this item. [L]

15 = Specified camber

Value >= 0. [L}

16 = Net area to total
area ratio

Value >= 0; 0 means
use program default value.

17 = Live load reduction
factor

Value >= 0; 0 means
use program determined value.

18 = Unbraced length ratio,
Major

Value >= 0; 0 means
use program determined value.

19 = Unbraced length ratio,
Minor

Value >= 0; 0 means
use program determined value.

20 = Unbraced length ratio,
Lateral Torsional Buckling

Value >= 0; 0 means
use program determined value.

21 = Effective length
factor, K1 Major

Value >= 0; 0 means
use program determined value.

22 = Effective length
factor, K1 Minor

Value >= 0; 0 means
use program determined value.

23 = Effective length
factor, K2 Major

Value >= 0; 0 means
use program determined value.

24 = Effective length
factor, K2 Minor

Value >= 0; 0 means
use program determined value.

25 = Effective length
factor, K Lateral Torsional Buckling

Value >= 0; 0 means
use program determined value.

26 = Moment coefficient,
Cm Major

Value >= 0; 0 means
use program determined value.

27 = Moment coefficient,
Cm Minor

Value >= 0; 0 means
use program determined value.

28 = Bending coefficient,
Cb

Value >= 0; 0 means
use program determined value.

29 = Nonsway moment factor,
B1 Major

Value >= 0; 0 means
use program determined value.

30 = Nonsway moment factor,
B1 Minor

Value >= 0; 0 means
use program determined value.

31 = Sway moment factor,
B2 Major

Value >= 0; 0 means
use program determined value.

32 = Sway moment factor,
B2 Minor

Value >= 0; 0 means
use program determined value.

33 = Reduce HSS thickness

0 = Program default

1 = No

2 = Yes

34 = HSS welding type

0 = Program default

1 = ERW

2 = SAW

35 = Yield stress, Fy

Value >= 0; 0 means
use program determined value. [F/L2]

36 = Expected to specified
Fy ratio, Ry

Value >= 0; 0 means
use program determined value.

37 = Compressive capacity,
Pnc

Value >= 0; 0 means
use program determined value. [F]

38 = Tensile capacity,
Pnt

Value >= 0; 0 means
use program determined value. [F]

39 = Major bending capacity,
Mn3

Value >= 0; 0 means
use program determined value. [FL]

40 = Minor bending capacity,
Mn2

Value >= 0; 0 means
use program determined value. [FL]

41 = Major shear capacity,
Vn2

Value >= 0; 0 means
use program determined value. [F]

42 = Minor shear capacity,
Vn3

Value >= 0; 0 means
use program determined value. [F]

43 = Demand/capacity ratio
limit

Value >= 0; 0 means use
program determined value.

ItemType

This is one of the following items in the eItemType
enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the
frame object specified by the Name item.

If this item is Group, the assignment is made to all
frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made
to all selected frame objects and the Name item is ignored.

## Remarks

This function sets the value of a steel design overwrite
item.

The function returns zero if the item is successfully
set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignOverwriteItemAISC360_16()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("AISC360-16")



   'set overwrite item

      ret = SapModel.DesignSteel.AISC360_16.SetOverwrite("8", 
 1, 7)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 20.0.0.

## See Also

[GetOverwrite](GetOverwrite_%7BAISC_360-16%7D.md)
