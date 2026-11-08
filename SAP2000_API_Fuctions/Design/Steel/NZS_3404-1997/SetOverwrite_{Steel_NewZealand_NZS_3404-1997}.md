# SetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.NewZealand\_NZS3404\_1997.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending on the value of the ItemType item.

Item

This is an integer between 1 and 46, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Steel type

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

19 = Unbraced length ratio, Minor

20 = Unbraced length ratio, Lateral Torsional Buckling

21 = Effective length factor, Ke Major Braced

22 = Effective length factor, Ke Minor Braced

23 = Effective length factor, Ke Major Sway

24 = Effective length factor, Ke Minor Sway

25 = Twist restraint factor for LTB (kt)

26 = lateral rotation restraint factor (kr)

27 = Load height factor for LTB (kl)

28 = Moment coefficient, Cm Major

29 = Moment coefficient, Cm Minor

30 = Moment modification factor, Alpha\_m

31 = Slender reduction factor, Alpha\_s

32 = Nonsway moment factor, Db Major

33 = Nonsway moment factor, Db Minor

34 = Sway moment factor, Ds Major

35 = Sway moment factor, Ds Minor

36 = Form factor, Kf

37 = Axial capacity correction factor, Kt

38 = Yield stress, Fy

39 = Compressive capacity, Nc

40 = Tensile capacity, Nt

41 = Major bending capacity, Ms33

42 = Minor bending capacity, Ms22

43 = Major bending capacity, Mb33

44 = Major shear capacity, Vu2

45 = Minor shear capacity, Vu3

46 = Demand/capacity ratio limit

Value

The value of the considered overwrite item.

1 = Framing type

0 = As specified in preferences

1 = Moment frame

2 = Braced frame

2 = Steel type

1 = Hot rolled

2 = Hot finished

3 = Cold form

4 = Stress relieved

5 = Lightly welded

6 = Heavily welded

3 = Consider deflection

0 = No

Any other value = Yes

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

  Value >= 0; 0 means no check for this item. [L]}

14 = Total camber limit, absolute

  Value >= 0; 0 means no check for this item. [L]

15 = Specified camber

  Value >= 0. [L}

16 = Net area to total area ratio

  Value >= 0; 0 means use program default value.

17 = Live load reduction factor

  Value >= 0; 0 means use program determined value.

18 = Unbraced length ratio, Major

  Value >= 0; 0 means use program determined value.

19 = Unbraced length ratio, Minor

  Value >= 0; 0 means use program determined value.

20 = Unbraced length ratio, Lateral Torsional Buckling

  Value >= 0; 0 means use program determined value.

21 = Effective length factor, Ke Major Braced

  Value >= 0; 0 means use program determined value.

22 = Effective length factor, Ke Minor Braced

  Value >= 0; 0 means use program determined value.

23 = Effective length factor, Ke Major Sway

  Value >= 0; 0 means use program determined value.

24 = Effective length factor, Ke Minor Sway

  Value >= 0; 0 means use program determined value.

25 = Twist restraint factor for LTB (kt)

  Value >= 0; 0 means use program determined value.

26 = Lateral rotation restraint factor (kr)

  Value >= 0; 0 means use program determined value.

27 = Load height factor for LTB (kl)

  Value >= 0; 0 means use program determined value.

28 = Moment coefficient, Cm Major

  Value >= 0; 0 means use program determined value.

29 = Moment coefficient, Cm Minor

  Value >= 0; 0 means use program determined value.

30 = Moment modification factor, Alpha\_m

  Value >= 0; 0 means use program determined value.

31 = Slender reduction factor, Alpha\_s

  Value >= 0; 0 means use program determined value.

32 = Nonsway moment factor, Db Major

  Value >= 0; 0 means use program determined value.

33 = Nonsway moment factor, Db Minor

  Value >= 0; 0 means use program determined value.

34 = Sway moment factor, Bs Major

  Value >= 0; 0 means use program determined value.

35 = Sway moment factor, Bs Minor

  Value >= 0; 0 means use program determined value.

36 = Form factor, Kf

  Value >= 0; 0 means use program determined value.

37 = Axial capacity correction factor, Kt

  Value >= 0; 0 means use program determined value.

38 = Yield stress, Fy

  Value >= 0; 0 means use program determined value. [F/L2]

39 = Compressive capacity, Nc

  Value >= 0; 0 means use program determined value. [F]

40 = Tensile capacity, Nt

  Value >= 0; 0 means use program determined value. [F]

41 = Major bending capacity, Ms33

  Value >= 0; 0 means use program determined value. [FL]

42 = Minor bending capacity, Ms22

  Value >= 0; 0 means use program determined value. [FL]

43 = Minor bending capacity, Mb33

  Value >= 0; 0 means use program determined value. [FL]

44 = Major shear capacity, Vu2

  Value >= 0; 0 means use program determined value. [F]

45 = Minor shear capacity, Vu3

  Value >= 0; 0 means use program determined value. [F]

46 = Demand/capacity ratio limit

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
Sub SetSteelDesignOverwriteItemAustralian_AS4100_1998 ()

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

      ret = SapModel.DesignSteel.SetCode("NZS 3404-1997")



   'set overwrite item

      ret = SapModel.DesignSteel.NewZealand_NZS3404_1997.SetOverwrite("8", 1, 7)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.1.

## See Also

[GetOverwrite](GetOverwrite_%7BSteel_NewZealand_NZS_3404-1997%7D.md)
