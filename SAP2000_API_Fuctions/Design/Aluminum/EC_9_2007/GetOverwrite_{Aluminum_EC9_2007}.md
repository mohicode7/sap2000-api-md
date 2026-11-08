# GetOverwrite

## Syntax

SapObject.SapModel.DesignAluminum.Eurocode\_9\_2007.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item
As Long, ByRef Value As Double, ByRef ProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with aluminum frame design
procedure.

Item

This is an integer between 1 and 29, inclusive, indicating
the overwrite item considered.

1 = Demand/capacity ratio
limit

2 = Live load reduction
factor

3 = Net area to total area
ratio

4 = Material buckling class

5
= Unbraced length ratio, y-y

6 = Unbraced length ratio,
z-z

7 = Unbraced length ratio,
LTB

8 = Effective length factor
sway, Ky

9 = Effective length factor
sway, Kz

10 = Effective length factor,
K LTB

11 = LTB fixed end

12 = Bending coefficient,
C1

13 = Bending coefficient,
C2

14 = Bending coefficient,
C3

15 = Torsional-flexural
buckling curve

16 = Buckling curve for
LTB

17 = Tensile capacity,
Nt,Rd

18 = Elastic torsional
buckling force, NcrT

19 = Elastic torsional-flexural
buckling force, NcrTF

20 = Compressive capacity,
NRk

21 = Bending capacity about
y-y axis, MyRk

22 = Bending capacity about
z-z axis, MzRk

23 = Elastic critical moment
for lateral-torsional buckling, Mcr

24 = Shear capacity along
z-z axis, Vz.Rd

25 = Shear capacity along
y-y axis, Vy.Rd

26 = Torsion capacity,
TRd

27 = Shear stress due to
warping torsion moment,  *τw*

28 = Warping coefficient,
kw

29 = Coordinate of load
application, za

Value

The value of the considered overwrite item.

1 = Demand/capacity ratio limit

      Value
>= 0; 0 means use program determined value.

2 = Live load reduction factor

      Value
>= 0; 0 means use a program determined value.

3 = Net area to total area ratio

      Value
>= 0; 0 means use program default value.

4 = Material buckling class

     0 = Program
default

     1 = Class
A

     2 = Class
B

5 = Unbraced length ratio, y-y

      Value
>= 0; 0 means use program determined value.

6 = Unbraced length ratio, z-z

      Value
>= 0; 0 means use program determined value.

7 = Unbraced length ratio, LTB

      Value
>= 0; 0 means use program determined value.

8 = Effective length factor, Ky

      Value
>= 0; 0 means use program determined value.

9 = Effective length factor, Kz

10 = Effective length factor, K LTB

       Value
>= 0; 0 means use program determined value.

11 = LTB Fixed End

       0
= Program default

       1
= Left

       2
= Right

12 = Bending coefficient, C1

       Value
>= 0; 0 means use program determined value.

13 = Bending coefficient, C2

       Value
>= 0; 0 means use program determined value.

14 = Bending coefficient, C3

       Value
>= 0; 0 means use program determined value.

15 = Torsional-flexural buckling curve

        0
= Program default

        1
= Curve 1

        2
= Curve 2

16 = Buckling curve for LTB

        0
= Program default

        1
= Curve 1

        2
= Curve 2

17 = Tensile capacity, Nt,Rd

       Value
>= 0; 0 means use program determined value. [F]

18 = Elastic torsional buckling force,
NcrT

       Value
>= 0; 0 means use program determined value. [F]

19 = Elastic torsional-flexural buckling
force, NcrTF

       Value
>= 0; 0 means use program determined value. [F]

20 = Compressive capacity, NRk

       Value
>= 0; 0 means use program determined value. [F]

  

21 = Bending capacity about y-y, MyRk

       Value
>= 0; 0 means use program determined value. [FL]

22 = Bending capacity about z-z, MzRk

       Value
>= 0; 0 means use program determined value. [FL]

23 = Elastic critical moment for lateral-torsional
buckling, Mcr

       Value
>= 0; 0 means use program determined value. [FL]

24 = Shear capacity along z-z axis, Vz.Rd

       Value
>= 0; 0 means use program determined value. [F]

25 = Shear capacity along y-y  axis,
Vy.Rd

       Value
>= 0; 0 means use program determined value. [F]

26 = Torsion capacity, TRd

       Value
>= 0; 0 means use program determined value. [FL]

27 = Shear Stress due to warping torsion
moment, *τw*

       Value
>= 0; 0 means use program determined value. [F/L2]

         

28 = Warping coefficient, kw
(used in Mcr calculation)

       0.5
=<Value =< 1; 0 means use program determined value which is defaulted
to 1.0.

29 = Coordinate of load application, za
(used in Mcr calculation)

ProgDet

If this item is True, the specified value is program
determined.

## Remarks

This function retrieves the value of an aluminum design
overwrite item.

The function returns zero if the item is successfully
retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAluminumDesignOverwriteItemEurocode_9_2007()

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
 
   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288) 



   'set aluminum design code

      ret = SapModel.DesignAluminum.SetCode("EN 
 1999:2007")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start aluminum design

      ret = SapModel.DesignAluminum.StartDesign



   'get overwrite item

      ret = SapModel.DesignAluminum.Eurocode_9_2007.GetOverwrite("8", 
 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.3.0

Changed references from major and minor axes to y-y
and z-z in version 23.4.0.

## See Also

[SetOverwrite](SetOverwrite_%7BAluminum_EC9_2007%7D.md)
