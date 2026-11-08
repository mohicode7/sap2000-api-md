# SetOverwrite

## Syntax

SapObject.SapModel.DesignSteel.ITALIAN\_UNI\_10011.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByValItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

Name

The name of a frame object with a steel frame design procedure.

Item

This is an integer between 1 and 26, inclusive, indicating the overwrite item considered.

1 = Framing type

2 = Axial load amplification(Omega)

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

19 = Unbraced length ratio, Minor Lateral TorsionalBuckling

20 = Effective length factor, Beta Major

21 = Effective length factor, Beta Minor

22 = Moment coefficient, Meq/Mmax Major

23 = Moment coefficient, Meq/Mmax Minor

24 = LTB moment coefficient (Omega1)

25 = Yield stress, Fy

26 = Demand/capacity ratio limit

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = Sway Frame

2 = NonSway Frame

2 = Axial load amplification(Omega)

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

19 = Unbraced length ratio, Minor Lateral TorsionalBuckling

Value >= 0; 0 means use program determined value.

20 = Effective length factor, Beta Major

Value >= 0; 0 means use program determined value.

21 = Effective length factor, Beta Minor

Value >= 0; 0 means use program determined value.

22 = Moment coefficient, Meq/Mmax Major

Value >= 0; 0 means use program determined value.

23 = Moment coefficient, Meq/Mmax Minor

Value >= 0; 0 means use program determined value.

24 = LTB moment coefficient (Omega1)

Value >= 0; 0 means use program determined value.

25 = Yield stress, Fy

Value >= 0; 0 means use program determined value. [F/L2]

26 = Demand/capacity ratio limit

Value >= 0; 0 means use program determined value.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjectst, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function sets the value of a steel design overwrite item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignOverwriteItemITALIAN_UNI_10011()

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

      ret = SapModel.DesignSteel.SetCode("ITALIAN UNI 10011")



   'set overwrite item

      ret = SapModel.DesignSteel.ITALIAN_UNI_10011.SetOverwrite("8", 1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetOverwrite](GetOverwrite_%7BSteel_Italian_UNI_10011%7D.md)
