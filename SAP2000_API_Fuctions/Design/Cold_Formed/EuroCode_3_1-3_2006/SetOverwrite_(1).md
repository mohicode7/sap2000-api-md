# **SetOverwrite**

## **Syntax**

SapObject.SapModel.Design<a id="ColdFormed"></a><a id="ColdFormed"></a>ColdFormed.EuroCold06.SetOverwrite

## **VB6 Procedure**

Function SetOverwrite(ByVal Name As String, ByVal Item
As Long, ByVal Value As Double, Optional ByVal ItemType As eItemType =
Object) As Long

## **Parameters**

Name

The name of a frame object with a cold-formed steel
frame design procedure.

Item

This is an integer between 1 and 34, inclusive, indicating
the overwrite item considered.

1 = Demand/capacity ratio limit

2 = Live load reduction factor

3 = Yield stress, Fy

4 = Net area to total area ratio

5 = Unbraced length ratio, y-y

6 = Unbraced length ratio, z-z

7 = Unbraced length ratio, LTB

8 = Effective length factor sway, K2y

9 = Effective length factor sway, K2z

10 = Effective length factor, K LTB

11 = Effective length factor braced, K1y

12 = Effective length factor braced, K1z

13 = Bending coefficient, C1

14 = Bending coefficient, C2

15 = Bending coefficient, C3

16 = Moment coefficient, kyy

17 = Moment coefficient, kzz

18 = Moment coefficient, kzy

19 = Moment coefficient, kyz

20 = Column buckling curve, y-y

21 = Column buckling curve, z-z

22 = Buckling curve for LTB

23 = Elastic torsional buckling force,
Ncr T

24 = Elastic torsional-flexural buckling
force, Ncr TF

25 = Compressive capacity, NRk

26 = Tensile capacity, Nt.Rd

27 = Bending capacity about the y-y axis,
MyRk

28 = Bending capacity about the z-z axis,
MzRk

29 = Shear capacity along z-z axis, Vbz.Rd

30 = Shear capacity along y-y axis, Vby.Rd

31 = Warping coefficient, kw

32 = Coordinate of load application, za

33 = Lateral-torsional buckling moment
capacity, Mcr

34 = Member buckling interaction equation

Value

The value of the considered overwrite item.

1 = Demand/capacity ratio limit

Value
>= 0; 0 means use program determined value.

2 = Live load reduction factor

Value
>= 0; 0 means use program determined value.

3 = Yield stress, Fy

Value
>= 0; 0 means use program determined value. [F/L2]

4 = Net area to total area ratio

Value
>= 0; 0 means use program default value.

5 = Unbraced length ratio, y-y

Value
>= 0; 0 means use program determined value.

6 = Unbraced length ratio, z-z

Value
>= 0; 0 means use program determined value.

7 = Unbraced length ratio, LTB

Value
>= 0; 0 means use program determined value.

8 = Effective length factor sway, K2y

Value >=
0; 0 means use program determined value.

9 = Effective length factor sway, K2z

Value
>= 0; 0 means use program determined value.

10 = Effective length factor, K LTB

Value
>= 0; 0 means use program determined value.

11 = Effective length factor braced, K1y

Value
>= 0; 0 means use program determined value.

12 = Effective length factor braced, K1z

Value
>= 0; 0 means use program determined value.

  

13 = Bending coefficient, C1

Value
>= 0; 0 means use program determined value.

14 = Bending coefficient, C2

Value >=
0; 0 means use program determined value.

15 = Bending coefficient, C3

Value
>= 0; 0 means use program determined value.

16 = Moment coefficient, kyy

Value
>= 0; 0 means use program determined value.

17 = Moment coefficient, kzz

Value
>= 0; 0 means use program determined value.

18 = Moment coefficient, kzy

Value
>= 0; 0 means use program determined value.

19 = Moment coefficient, kyz

Value
>= 0; 0 means use program determined value.

  

20 = Column buckling curve, y-y

0
= Program default

1
= a0

2
= a

3
= b

4
= c

5
= d

21 = Column buckling curve, z-z

0
= Program default

1
= a0

2
= a

3
= b

4
= c

5
= d

22 = Buckling curve for LTB

0
= Program default

1
= a0

2
= a

3
= b

4
= c

5
= d

23 = Elastic torsional buckling force,
Ncr T

Value
>= 0; 0 means use program determined value. [F]

24 = Elastic torsional-flexural buckling
force, Ncr TF

Value
>= 0; 0 means use program determined value. [F]

25 = Compressive capacity, NRk

Value
>= 0; 0 means use program determined value. [F]

26 = Tensile capacity, Nt.Rd

Value
>= 0; 0 means use program determined value. [F]

27 = Bending capacity about the y-y axis,
MyRk

Value
>= 0; 0 means use program determined value. [FL]

28 = Bending capacity about the z-z axis,
MzRk

Value
>= 0; 0 means use program determined value. [FL]

  

29 = Shear capacity along z-z axix, Vbz.Rd

Value
>= 0; 0 means use program determined value. [F]

30 = Shear capacity along y-y axis, Vby.Rd

Value
>= 0; 0 means use program determined value. [F]

         

31 = Warping coefficient, kw
(used in Mcr calculation)

0.5
=<Value =< 1; 0 means use program determined value which is defaulted
to 1.0.

32 = Coordinate of load application, za
(used in Mcr calculation)

33 = Lateral-torsional buckling moment
capacity, Mcr

Value
>= 0; 0 means use program determined value. [FL]

34 = Member buckling interaction equations

1
= Equations 6.61 and 6.62 in EN 1993-1-1:2005

2 = Equation 6.36 in EN 1993-1-3:2006 for each
axis of bending separately

3
= Equation 6.36 in EN 1993-1-3:2006 for both axes of bending simultaneously

4
= Equation 6.61 and 6.62 in EN 1993-1-1:2005 and Equation 6.36 in EN 1993-1-3:2006
for each axis of bending separately

5
= Equation 6.61 and 6.62 in EN 1993-1-1:2005 and Equation 6.36 in EN 1993-1-3:2006
for both axes of bending simultaneously

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

## **Remarks**

This function sets the value of a cold-formed steel
design overwrite item.

The function returns zero if the item is successfully
set; otherwise, it returns a nonzero value.

## **VBA Example**



```
Sub SetSteelDesignOverwriteItemEuroCold06()

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

      ret = SapModel.DesignColdFormed.SetCode("Eurocode 
 3 1-3 2006")



   'set overwrite item

      ret = SapModel.DesignColdFormed.EuroCold06.SetOverwrite("8", 
 1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## **Release Notes**

Initial release in version 23.0.0.

Added item 34 in version 23.4.0

Changed references from major and minor axes to y-y
and z-z in version 23.4.0.

## **See Also**

[GetOverwrite](GetOverwrite_%281%29.md)
