# SetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.INDIAN\_IS\_800\_1998.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByValItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending on the value of the ItemType item.

Item

This is an integer between 1 and 34, inclusive, indicating the overwrite item considered.

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

18 = Unbraced length ratio, Minor

19 = Unbraced length ratio, Lateral Torsional Buckling

20 = Effective length factor, K Major

21 = Effective length factor, K Minor

22 = Effective length factor, K Lateral Torsional Buckling

23 = Moment coefficient, Cm Major

24 = Moment coefficient, Cm Minor

25 = Yield stress, Fy

26 = Allowable compressive stress, Sigma\_ac

27 = Allowable tensile stress, Sigma\_at

28 = Allowable major bending stress, Sigma\_bc33

29 = Allowable minor bending stress, Sigma\_bc22

30 = Major average shear stress, Tau\_va2

31 = Minor average shear stress, Tau\_va3

32 = Maximum elastic shear stress, Tau\_vm

33 = Allowable effective stress, Sigma\_e

34 = Demand/capacity ratio limit

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = Sway Frame

2 = Nonsway Frame

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

18 = Unbraced length ratio, Minor

Value >= 0; 0 means use program determined value.

19 = Unbraced length ratio, Lateral Torsional Buckling

Value >= 0; 0 means use program determined value.

20 = Effective length factor, K Major

Value >= 0; 0 means use program determined value.

21 = Effective length factor, K Minor

Value >= 0; 0 means use program determined value.

22 = Effective length factor, K Lateral Torsional Buckling

Value >= 0; 0 means use program determined value.

23 = Moment coefficient, Cm Major

Value >= 0; 0 means use program determined value.

24 = Moment coefficient, Cm Minor

Value >= 0; 0 means use program determined value.

25 = Yield stress, Fy

Value >= 0; 0 means use program determined value. [F/L2]

26 = Allowable compressive stress, Sigma\_ac

Value >= 0; 0 means use program determined value. [F/L2]

27 = Allowable tensile stress, Sigma\_at

Value >= 0; 0 means use program determined value. [F/L2]

28 = Allowable major bending stress, Sigma\_bc33

Value >= 0; 0 means use program determined value. [F/L2]

29 = Allowable minor bending stress, Sigma\_bc22

Value >= 0; 0 means use program determined value. [F/L2]

30 = Major average shear stress, Tau\_va2

Value >= 0; 0 means use program determined value. [F/L2]

31 = Minor average shear stress, Tau\_va3

Value >= 0; 0 means use program determined value. [F/L2]

32 = Maximum elastic shear stress, Tau\_vm

Value >= 0; 0 means use program determined value. [F/L2]

33 = Allowable effective stress, Sigma\_e

Value >= 0; 0 means use program determined value. [F/L2]

34 = Demand/capacity ratio limit

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
Sub SetSteelDesignOverwriteItemINDIAN_IS_800_1998()

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

      ret = SapModel.DesignSteel.SetCode("Indian IS:800-1998")



   'set overwrite item

      ret = SapModel.DesignSteel.INDIAN_IS_800_1998.SetOverwrite("8", 1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetOverwrite](GetOverwrite_%7BSteel_Indian_IS_800-1998%7D.md)
