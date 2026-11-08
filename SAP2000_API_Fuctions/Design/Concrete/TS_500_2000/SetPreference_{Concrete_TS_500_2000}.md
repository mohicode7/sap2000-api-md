# SetPreference

SapObject.SapModel.DesignConcrete.TS\_500\_2000.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value
As Double) As Long

## Parameters

Name

The name of an existing frame object or group, depending
on the value of the ItemType item.

Item

This is an integer between 1 and 12, inclusive, indicating
the overwrite item considered.

1 = Number of interaction
curves

2 = Number of interaction
points

3 = Consider maximum eccentricity

4 = Seismic zone

5 = Gamma steel

6 = Gamma concrete

7 = Gamma concrete shear

8 = Pattern live load
factor

9 = Utilization factor
limit

10 = Multi-response case
design

11
= Consider torsion

12 = Design for B/C capacity
ratio

Value

The value of the considered preference item.

1 = Number of interaction
curves

Value >= 4 and divisible
by 4

2 = Number of interaction
points

Value >= 5 and odd

3 = Consider maximum eccentricity

0 = No

Any other value = Yes

4 = Seismic zone

1 = Zone 1

2 = Zone 2

3 = Zone 3

4 = Zone 4

5 = Gamma steel

Value > 0

6 = Gamma concrete

Value > 0

7 = Gamma concrete shear

Value > 0

8 = Pattern live load
factor

Value >= 0

9 = Utilization factor
limit

Value >= 0.

10 = Multi-response case
design

1 = Envelopes

2 = Step-by-Step

3 = Last Step

4 = Envelopes -- All

5 = Step-by-Step -- All

11
= Consider torsion

0 = No

Any other value = Yes

12 = Design for B/C capacity
ratio

0 = No

Any other value = Yes

## Remarks

This function sets the value of a concrete design preference
item.

The function returns zero if the item is successfully
set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignPreferenceItemTS_500_2000()

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



   'create new concrete frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", 
 "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288, True, "R1", "R1")



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("TS 
 500-2000")



   'set overwrite item

      ret = SapModel.DesignConcrete.TS_500_2000.SetPreference(2, 
 9)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

Added items 11-12 in version 23.4.0

## See Also

[GetPreference](GetPreference_%7BConcrete_TS_500-2000%7D.md)
