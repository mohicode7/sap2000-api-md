# GetOverwrite

SapObject.SapModel.DesignConcrete.TS\_500\_2000.GetOverwrite

## VB6 Procedure

Function GetOverwrite(ByVal Name As String, ByVal Item
As Long, ByRef Value As Double, ByRefProgDet As Boolean) As Long

## Parameters

Name

The name of a frame object with a concrete frame design
procedure.

Item

This is an integer between 1 and 14, inclusive, indicating
the overwrite item considered.

1 = Framing type

2 = Live load reduction
factor

3 = Unbraced length ratio,
Major

4 = Unbraced length ratio,
Minor

5 = Effective length factor,
k Major

6 = Effective length factor,
k Minor

7 = Moment coefficient,
Cm Major

8 = Moment coefficient,
Cm Minor

9 = Non-sway moment factor,
Bns Major

10 = Non-sway moment factor,
Bns Minor

11 = Sway moment factor,
Bs Major

12 = Sway
moment factor, Bs Minor

13
= Consider torsion (beam only) (beam only)

14 = Concrete cover for
closed stirrups (beam only)

Value

The value of the considered preference item.

1 = Framing type

0 = Program Default

1 = High Ductile

2 = Nominal Ductile

3 = Ordinary

4 = Non-sway

2 = Live load reduction
factor

Value >= 0; 0 means
use program determined value

3 = Unbraced length ratio,
Major

Value >= 0; 0 means
use program determined value

4 = Unbraced length ratio,
Minor

Value >= 0; 0 means
use program determined value

5 = Effective length factor,
k Major

Value >= 0; 0 means
use program determined value

6 =  Effective length
factor, k Minor

Value >= 0; 0 means
use program determined value

7 = Moment coefficient,
Cm Major

Value >= 0; 0 means
use program determined value

8 = Moment coefficient,
Cm Minor

Value >= 0; 0 means
use program determined value

9 = Non-sway moment factor,
Bns Major

Value >= 0; 0 means
use program determined value

10 = Non-sway moment factor,
Bns Minor

Value >= 0; 0 means
use program determined value

11 = Sway moment factor,
Bs Major

Value >= 0; 0 means
use program determined value

12 = Sway moment factor,
Bs Minor

Value >= 0; 0 means
use program determined value

13 = Consider torsion

0 = No

Any other value = Yes

14 = Concrete cover for
closed stirrups

Value >= 0; 0 means
use program determined value.

## ProgDet

If this Item is true, the specified value is program
determined.

## Remarks

This function sets the value of a concrete design overwrite
item.

The function returns zero if the item is successfully
set; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignOverwriteItemTS_500_2000()

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

      ret = SapModel.DesignConcrete.SetCode("TS 
 500-2000")



   'run analysis
      ret = SapModel.File.Save("C:\SapAPI\x.sdb")
      ret = SapModel.Analyze.RunAnalysis
 
   'start concrete design
      ret = SapModel.DesignConcrete.StartDesign
 
   'get overwrite item

      ret = SapModel.DesignConcrete.TS_500_2000.GetOverwrite("8", 
 1, Value, ProgDet)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

Added items 13-14 in version 23.4.0

## See Also

[SetOverwrite](SetOverwrite_%7BTS_500_2000%7D.md)
