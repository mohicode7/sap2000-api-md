# SetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.SP\_16\_13330\_2011.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending on the value of the ItemType item.

Item

This is an integer between 1 and 37, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Section Class

3 = Service Factor, GammaC

4 = Service Factor, GammaU

5 = Factor for L-Shapes , GammaC1

6 = Column buckling curve major

7 = Column buckling curve minor

8 = Is rolled section?

9 = Is beam top loaded?

10 = Consider Deflection

11 = Deflection check type

12 = DL deflection limit, L/Value

13 = SDL+LL deflection limit, L/Value

14 = LL deflection limit, L/Value

15 = Total deflection limit, L/Value

16 = Total-camber limit, L/Value

17 = DL deflection limit, absolute

18 = SDL+LL deflection limit, absolute

19 = LL deflection limit, absolute

20 = Total deflection limit, absolute

21 = Total-camber deflection limit, absolute

22 = Specified camber

23 = Net area to total area ratio

24 = Live load reduction factor

25 = Unbraced length ratio, Major

26 = Unbraced length ratio, Minor

27 = Unbraced length ratio, LTB

28 = Effective length factor, K1 Major

29 = Effective length factor, K1 Minor

30 = Effective length factor, K2 Major

31 = Effective length factor, K2 Minor

32 = Effective length factor, LTB

33 = Characteristic strength, Ryn

34 = Design yield strength, Ry

35 = Design fracture strength, Ru

36 = Design shear strength, Rs

37 = Demand/capacity ratio, limit

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = Moment Frame

2 = Braced Frame

2 = Section Class

0 = Program Determined

1 = Class 1

2 = Class 2

3 = Class 3

4 = Class 4

3 = Service Factor, GammaC

Value >= 0; 0 means use program determined value.

4 = Service Factor, GammaU

Value >= 0; 0 means use program determined value.

5 = Factor for L-Shapes , GammaC1

Value >= 0; 0 means use program determined value.

6 = Column buckling curve major

0 = Program Determined

1 = a

2 = b

3 = c

7 = Column buckling curve minor

0 = Program Determined

1 = a

2 = b

3 = c

8 = Is rolled section?

0 = Program Determined

1 = No

2 = Yes

9 = Is beam top loaded?

0 = Program Determined

1 = No

2 = Yes

10 = Consider Deflection

0 = Program Determined

1 = No

2 = Yes

11 = Deflection check type

0 = Program Determined

1 = Ratio

2 = Absolute

3 = Both

12 = DL deflection limit, L/Value

  Value >= 0; 0 means no check for this item.

13 = SDL+LL deflection limit, L/Value

  Value >= 0; 0 means no check for this item.

14 = LL deflection limit, L/Value

  Value >= 0; 0 means no check for this item.

15 = Total deflection limit, L/Value

  Value >= 0; 0 means no check for this item.

16 = Total-camber limit, L/Value

  Value >= 0; 0 means no check for this item.

17 = DL deflection limit, absolute

  Value >= 0; 0 means no check for this item. [L]

18 = SDL+LL deflection limit, absolute

  Value >= 0; 0 means no check for this item. [L]

19 = LL deflection limit, absolute

  Value >= 0; 0 means no check for this item. [L]

20 = Total deflection limit, absolute

  Value >= 0; 0 means no check for this item. [L]

21 = Total-camber deflection limit, absolute

  Value >= 0; 0 means no check for this item. [L]

22 = Specified camber

  Value >= 0. [L]

23 = Net area to total area ratio

  Value >= 0; 0 means use program default value.

24 = Live load reduction factor

  Value >= 0; 0 means use program determined value.

25 = Unbraced length ratio, Major

  Value >= 0; 0 means use program determined value.

26 = Unbraced length ratio, Minor

  Value >= 0; 0 means use program determined value.

27 = Unbraced length ratio, LTB

  Value >= 0; 0 means use program determined value.

28 = Effective length factor, K1 Major

  Value >= 0; 0 means use program determined value.

29 = Effective length factor, K1 Minor

  Value >= 0; 0 means use program determined value.

30 = Effective length factor, K2 Major

  Value >= 0; 0 means use program determined value.

31 = Effective length factor, K2 Minor

  Value >= 0; 0 means use program determined value.

32 = Effective length factor, LTB

  Value >= 0; 0 means use program determined value.

33 = Characteristic strength, Ryn

  Value >= 0; 0 means use program determined value. [F/L2]

34 = Design yield strength, Ry

  Value >= 0; 0 means use program determined value. [F/L2]

35 = Design fracture strength, Ru

Value >= 0; 0 means use program determined value. [F/L2]

36 = Design shear strength, Rs

Value >= 0; 0 means use program determined value. [F/L2]

37 = Demand/capacity ratio, limit

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

The function returns zero if the item is successfully set; otherwise, it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignOverwriteItemSP16_13330_2011()

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

      ret = SapModel.DesignSteel.SetCode("SP 16.13330.2011")



   'set overwrite item

      ret = SapModel.DesignSteel.SP16_13330_2011.SetOverwrite("8", 1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v20.1.0.

## See Also

[GetOverwrite](GetOverwrite_%7BSP63_16-13330-2011%7D.md)
