# GetOverwrite

## Syntax

SapObject.SapModel.DesignConcrete.Indian\_IS\_456\_2000.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item
As Long, ByRef Value As Double, ByRefProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a concrete frame design
procedure.

Item

This is an integer between 1 and 22, inclusive, indicating
the overwrite item considered.

1 = Framing type (beam
and column)

2 = Live load reduction
factor (beam and column)

3 = Unbraced length ratio,
Major (beam and column)

4 = Unbraced length ratio,
Minor (beam and column)

5 = Effective length factor,
K Major (column only)

6 = Effective length factor,
K Minor (column only)

7 = Moment coefficient,
Cm Major (not used)

8 = Moment coefficient,
Cm Minor (not used)

9 = Non-sway moment factor,
Dns Major (not used)

10 = Non-sway moment factor,
Dns Minor (not used)

11 = Sway moment factor,
Ds Major (not used)

12 = Sway moment factor, Ds
Minor (not used)

13 = Top rebar area of a beam
at the left end (I-end) (beam only)

14 = Bottom rebar area of
a beam at the left end (I-end) (beam
only)

15 = Top rebar area of a beam
at the right end (J-end) (beam only)

16 = Bottom rebar area of
a beam at the right end (J-end) (beam only)

17 = Consider torsion (not
used)

18 = Concrete cover for
closed stirrups (beam only)

19 = Effective length factor
braced, K Major braced (column only)

20 = Effective length factor
braced, K Minor braced (column only)

21 = Q factor in global
X direction (column only)

22 = Q factor in global
Y direction (column only)

Value

The value of the considered overwrite item.

1 = Framing type

0 = Program Default

1 = Ductile

2 = Ordinary

3 = Non-sway

2 = Live load reduction
factor

Value >= 0; 0 means
use program determined value.

3 = Unbraced length ratio,
Major

Value >= 0; 0 means
use program determined value.

4 = Unbraced length ratio,
Minor

Value >= 0; 0 means
use program determined value.

5 = Effective length factor,
K Major

Value >= 0; 0 means
use program determined value.

6 = Effective length factor,
K Minor

Value >= 0; 0 means
use program determined value.

7 = Moment coefficient,
Cm Major

Value >= 0; 0 means
use program determined value.

8 = Moment coefficient,
Cm Minor

Value >= 0; 0 means
use program determined value.

9 = Non-sway moment factor,
Dns Major

Value >= 0; 0 means
use program determined value.

10 = Non-sway moment factor,
Dns Minor

Value >= 0; 0 means
use program determined value.

11 = Sway moment factor,
Ds Major

Value >= 0; 0 means
use program determined value.

12 = Sway moment factor,
Ds Minor

Value >= 0; 0 means
use program determined value.

13
= Top rebar area of a beam at the left end (I-end)

        Value
>= 0; 0 means use program determined value.

14
= Bottom rebar area of a beam at the left end (I-end)

        Value
>= 0; 0 means use program determined value.

15
= Top rebar area of a beam at the right end (J-end)

        Value
>= 0; 0 means use program determined value.

16
= Bottom rebar area of a beam at the right end (J-end)

        Value
>= 0; 0 means use program determined value.

17 = Consider torsion (not
used)

0 = No

Any other value = Yes

18 = Concrete cover for
closed stirrups

Value >= 0; 0 means
use program determined value.

19 = Effective length factor
braced, K Major braced

Value >= 0; 0 means
use program determined value.

20 = Effective length factor
braced, K Minor braced

Value >= 0; 0 means
use program determined value.

21 = Q factor in global
X direction

Value >= 0; 0 means
use program determined value.

22 = Q factor in global
Y direction

Value >= 0; 0 means
use program determined value.

ProgDet

If this item is True, the specified value is program
determined.

## Remarks

This function retrieves the value of a concrete design
overwrite item.

The function returns zero if the item is successfully
retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignOverwriteItemIndian_IS_456_2000()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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



   'create new concrete frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", 
 "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288, True, "R1", "R1")



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("Indian 
 IS 456-2000")



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'start concrete design

      ret = SapModel.DesignConcrete.StartDesign



   'get overwrite item

      ret = SapModel.DesignConcrete.Indian_IS_456_2000.GetOverwrite("8", 
 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Added Items 13, 14, 15, and 16 in Version 14.0.0.

Changed description for items 7 through 12 and added
items 17-22 in v23.4.0.

Changed description of item 17 to “not used” in version
24.2.0

## See Also

[SetOverwrite](SetOverwrite_%7BConcrete_Indian_IS_456_2000%7D.md)
