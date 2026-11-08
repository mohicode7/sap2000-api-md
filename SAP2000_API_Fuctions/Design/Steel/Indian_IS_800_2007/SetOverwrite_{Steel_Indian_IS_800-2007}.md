# SetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.Indian\_IS\_800\_2007.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending on the value of the ItemType item.

Item

This is an integer between 1 and 44, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Section class

3 = Column buckling curve (z-z)

4 = Column buckling curve (y-y)

5 = Is rolled section

6 = Consider deflection

7 = Deflection check type

8 = DL deflection limit, L/Value

9 = SDL + LL deflection limit, L/Value

10 = LL deflection limit, L/Value

11 = Total load deflection limit, L/Value

12 = Total camber limit, L/Value

13 = DL deflection limit, absolute

14 = SDL + LL deflection limit, absolute

15 = LL deflection limit, absolute

16 = Total load deflection limit, absolute

17 = Total camber limit, absolute

18 = Specified camber

19 = Net area to total area ratio

20 = Live load reduction factor

21 = Unbraced length ratio, Major

22 = Unbraced length ratio, Minor

23 = Unbraced length ratio, Lateral Torsional Buckling

24 = Effective length factor Braced, K1 Major

25 = Effective length factor Braced, K1 Minor

26 = Effective length factor Sway, K2 Major

27 = Effective length factor Sway, K2 Minor

28 = Effective length factor, K Lateral Torsional Buckling

29 = Bending coefficient, C1

30 = Uniform moment factor, Cmz

31 = Uniform moment factor, Cmy

32 = Uniform moment factor, CmLT

33 = Moment coefficient, kz

34 = Moment coefficient, ky

35 = Moment coefficient, k\_LT

36 = Yield stress, Fy

37 = Compressive capacity, Pd

38 = Tensile capacity, Td

39 = Major bending capacity, Mdz

40 = Minor bending capacity, Mdy

41 = Critical buckling moment, Mcr

42 = Major shear capacity, Vdy

43 = Minor shear capacity, Vdz

44 = Demand/capacity ratio limit

Value

The value of the considered overwrite item.

1 = Framing type

0 = As specified in preferences

1 = SMF

2 = OMF

3 = SCBF

4 = OCBF

5 = EBF

6 = Secondary

2 = Section class

1 = Class 1 (Plastic)

2 = Class 2 (Compact)

3 = Class 3 (Semicompact)

4 = Class 4 (Slender)

3 = Column buckling curve (z-z)

1 = a

2 = b

3 = c

4 = d

4 = Column buckling curve (y-y)

1 = a

2 = b

3 = c

4 = d

5 = Is rolled section

      Value >= 0; 0 means use program default value

6 = Consider deflection

0 = Program Determined

1 = No

2 = Yes

7 = Deflection check type

0 = Program default

1 = Ratio

2 = Absolute

3 = Both

8 = DL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

9 = SDL + LL deflection limit, L/Value

Value >= 0; 0 means no check for this item.

10 = LL deflection limit, L/Value

  Value >= 0; 0 means no check for this item.

11 = Total load deflection limit, L/Value

  Value >= 0; 0 means no check for this item.

12 = Total camber limit, L/Value

  Value >= 0; 0 means no check for this item.

13 = DL deflection limit, absolute

  Value >= 0; 0 means no check for this item. [L]

14 = SDL + LL deflection limit, absolute

  Value >= 0; 0 means no check for this item. [L]

15 = LL deflection limit, absolute

  Value >= 0; 0 means no check for this item. [L]

16 = Total load deflection limit, absolute

  Value >= 0; 0 means no check for this item. [L]

17 = Total camber limit, absolute

  Value >= 0; 0 means no check for this item. [L]

18 = Specified camber

  Value >= 0. [L]

19 = Net area to total area ratio

  Value >= 0; 0 means use program default value.

20 = Live load reduction factor

  Value >= 0; 0 means use program determined value.

21 = Unbraced length ratio, Major

  Value >= 0; 0 means use program determined value.

22 = Unbraced length ratio, Minor

  Value >= 0; 0 means use program determined value.

23 = Unbraced length ratio, Lateral Torsional Buckling

  Value >= 0; 0 means use program determined value.

24 = Effective length factor Braced, K1 Major

  Value >= 0; 0 means use program determined value.

25 = Effective length factor Braced, K1 Minor

  Value >= 0; 0 means use program determined value.

26 = Effective length factor Sway, K2 Major

  Value >= 0; 0 means use program determined value.

27 = Effective length Sway factor, K2 Minor

  Value >= 0; 0 means use program determined value.

28 = Effective length factor, K Lateral Torsional Buckling

  Value >= 0; 0 means use program determined value.

29 = Bending coefficient, C1

  Value >= 0; 0 means use program determined value.

30 = Uniform moment factor, Cmz

  Value >= 0; 0 means use program determined value.

31 = Uniform moment factor, Cmy

  Value >= 0; 0 means use program determined value.

32 = Uniform moment factor, CmLT

  Value >= 0; 0 means use program determined value.

33 = Moment coefficient, kz

  Value >= 0; 0 means use program determined value.

34 = Moment coefficient, ky

  Value >= 0; 0 means use program determined value.

35 = Moment coefficient, k\_LT

  Value >= 0; 0 means use program determined value.

36 = Yield stress, Fy

  Value >= 0; 0 means use program determined value. [F/L2]

37 = Compressive capacity, Pd

  Value >= 0; 0 means use program determined value. [F]

38 = Tensile capacity, Td

  Value >= 0; 0 means use program determined value. [F]

39 = Major bending capacity, Mdz

  Value >= 0; 0 means use program determined value. [F-L]

40 = Minor bending capacity, Mdy

  Value >= 0; 0 means use program determined value. [F-L]

41 = Critical buckling moment, Mcr

  Value >= 0; 0 means use program determined value. [F-L]

42 = Major shear capacity, Vdy

  Value >= 0; 0 means use program determined value. [F]

43 = Minor shear capacity, Vdz

  Value >= 0; 0 means use program determined value. [F]

44 = Demand/capacity ratio limit

  Value >= 0; 0 means use program determined value.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects and the Name item is ignored.

## Remarks

This function sets the value of a steel design overwrite item.

The function returns zero if the item is successfully set, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignOverwriteItemINDIAN_IS_800_2007()

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

      ret = SapModel.DesignSteel.SetCode("Indian IS:800-2007")



   'set overwrite item

      ret = SapModel.DesignSteel.INDIAN_IS_800_2007.SetOverwrite("8", 1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.1.

## See Also

[GetOverwrite](GetOverwrite_%7BSteel_Indian_IS_800-2007%7D.md)
