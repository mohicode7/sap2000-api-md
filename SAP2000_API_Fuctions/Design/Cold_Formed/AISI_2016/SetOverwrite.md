# SetOverwrite

## Syntax

SapObject.SapModel.DesignColdFormed.AISI\_ASD96.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item
As Long, ByVal Value As Double, Optional ByVal ItemType As eItemType =
Object) As Long

## Parameters

Name

The name of an existing frame object or group depending
on the value of the ItemType item.

Item

This is an integer between 1 and 33, inclusive, indicating
the overwrite item considered.

1 = Framing type

2 = Demand/capacity ratio limit

3 = Live load reduction factor

4 = Unbraced length ratio, Major

5 = Unbraced length ratio, Minor

6 = Unbraced length ratio, Lateral Torsional
Buckling

7 = Unbraced length ratio, Distortional Buckling

8 = Effective length factor, K Major

9 = Effective length factor, K Minor

10 = Effective length factor, K Lateral Torsional
Buckling

11 = Effective length factor, K Distortional
Buckling

12 = Rotational stiffness, kPhi (Distortional
Buckling)

13 = Moment coefficient, Cm Major

14 = Moment coefficient, Cm Minor

15 = Moment coefficient, Ctf Major

16 = Moment coefficient, Ctf Minor

17 = Bending coefficient, Cb Major

18 = Bending coefficient, Cb Minor

19 = Moment factor, Alpha Major

20 = Moment factor, Alpha Minor

21 = Nonsway moment factor, B1 Major

22 = Nonsway moment factor, B1 Minor

23 = Sway moment factor, B2 Major

24 = Sway moment factor, B2 Minor

25 = Through fastened to deck

26 = Fastener eccentricity, a/b

27 = Hole diameter on web

28 = Hole length on web

29 = Hole spacing on web

30 = Asymmetric integral ratio, Major

31 = Asymmetric integral ratio, Minor

32 = Nominal bending capacity for lateral torsional
buckling, Mn33

33 = Nominal bending capacity for lateral torsional
buckling, Mn22

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program
Default

1 = Moment
Frame

2 = Braced
Frame

2 = Demand/capacity ratio limit

Value >=
0; 0 means use program determined value

3 = Live load reduction factor

Value >=
0; 0 means use a program determined value.

4 = Unbraced length ratio, Major

Value >=
0; 0 means use program determined value.

5 = Unbraced length ratio, Minor

Value >=
0; 0 means use program determined value.

6 = Unbraced length ratio, Lateral Torsional
Buckling

Value >=
0; 0 means use program determined value.

7 = Unbraced length ratio, Distortional
Buckling

Value >=
0; 0 means use program determined value.

8 = Effective length factor, K Major

Value >=
0; 0 means use program determined value.

9 = Effective length factor, K Minor

Value >=
0; 0 means use program determined value.

10 = Effective length factor, K Lateral
Torsional Buckling

Value >=
0; 0 means use program determined value.

11 = Effective length factor, K Distortional
Buckling

Value >=
0; 0 means use program determined value.

12 = Rotational stiffness, kPhi (Distortional
Buckling)

Value >=
0; 0 means use program determined value. [F]

13 = Moment coefficient, Cm Major

Value >=
0; 0 means use program determined value.

14 = Moment coefficient, Cm Minor

Value >=
0; 0 means use program determined value.

15 = Moment coefficient, Ctf Major

Value >=
0; 0 means use program determined value.

16 = Moment coefficient, Ctf Minor

Value >=
0; 0 means use program determined value.

17 = Bending coefficient, Cb Major

Value >=
0; 0 means use program determined value.

18 = Bending coefficient, Cb Minor

Value >=
0; 0 means use program determined value.

19 = Moment factor, Alpha Major

Value >=
0; 0 means use program determined value.

20 = Moment factor, Alpha Minor

Value >=
0; 0 means use program determined value.

21 = Nonsway moment factor, B1 Major

Value >=
0; 0 means use program determined value.

22 = Nonsway moment factor, B1 Minor

Value >=
0; 0 means use program determined value.

23 = Sway moment factor, B2 Major

Value >=
0; 0 means use program determined value.

24 = Sway moment factor, B2 Minor

Value >=
0; 0 means use program determined value.

25 = Through fastened to deck

0 = No.

Any other
value = Yes.

26 = Fastener eccentricity, a/b

Value >=
0; 0 means use program determined value. [L]

  

27 = Hole diameter on web

Value >=
0; 0 means use program determined value. [L]

28 = Hole length on web

Value >=
hole diameter on web; 0 means use program determined value. [L]

29 = Hole spacing on web

Value >=
hole length on web; 0 means use program determined value. [L]

30 = Asymmetric integral ratio, Major

Value >=
0; 0 means use program determined value.

31 = Asymmetric integral ratio, Minor

Value >=
0; 0 means use program determined value.

32 = Nominal bending capacity for lateral
torsional buckling, Mn33

Value >=
0; 0 means use program determined value. [FL]

33 = Nominal bending capacity for lateral
torsional buckling, Mn22

Value >=
0; 0 means use program determined value. [FL]

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
to all selected frame objects, and the Name item is ignored.

## Remarks

This function sets the value of a cold formed design
overwrite item.

The function returns zero if the item is successfully
set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetColdFormedDesignOverwriteItemAISI_16()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, 
 MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section property

      ret = SapModel.PropFrame.SetColdC("CdC", 
 Name , 9, 3, 0.06, 0.25, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 72, 2, 144, True, "CdC", "CdC")



   'set cold formed design code

      ret = SapModel.DesignColdFormed.SetCode("AISI-16")



   'set overwrite item

      ret = SapModel.DesignColdFormed.AISI_16.SetOverwrite("8", 
 1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 22.2.0.

## See Also

[GetOverwrite](GetOverwrite.md)
