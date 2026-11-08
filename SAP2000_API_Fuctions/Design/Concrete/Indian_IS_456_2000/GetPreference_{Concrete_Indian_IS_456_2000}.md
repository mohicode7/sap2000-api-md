# GetPreference

## Syntax

SapObject.SapModel.DesignConcrete.Indian\_IS\_456\_2000.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value
As Double) As Long

## Parameters

Item

This is an integer between 1 and 12, inclusive, indicating
the preference item considered.

1 = Number of interaction
curves

2 = Number of interaction
points

3 = Consider minimum eccentricity

4 = Gamma steel

5 = Gamma concrete

6 = Pattern live load
factor

7 = Utilization factor
limit

8 = Multi-response case design

9 = Consider torsion (not
used)

10 = Consider additional
moment

11 = Consider P-Delta done

12 = Design for B/C capacity
ratio

Value

The value of the considered preference item.

1 = Number of interaction
curves

Value >= 4 and devisable
by 4

2 = Number of interaction
points

Value >= 5 and odd

3 = Consider minimum eccentricity

0 = No

Any other value = Yes

4 = Gamma steel

Value > 0

5 = Gamma concrete

Value > 0

6 = Pattern live load
factor

Value >= 0

7 = Utilization factor
limit

Value > 0

8 = Multi-response case
design

1 = Envelopes

2 = Step-by-step

3 = Last step

4 = Envelopes -- All

5 = Step-by-step -- All

9
= Consider torsion (not used)

0 = No

Any other value = Yes

10 = Consider additional
moment

0 = No

Any other value = Yes

11 = Consider P-Delta done

0 = No

Any other value = Yes

12 = Design for B/C capacity
ratio

0 = No

Any other value = Yes

## Remarks

This function retrieves the value of a concrete design
preference item.

The function returns zero if the item is successfully
retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignPreferenceItemIndian_IS_456_2000()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value As Double



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



   'get preference item

      ret = SapModel.DesignConcrete.Indian_IS_456_2000.GetPreference(2, 
 Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed Time history design item to Multi-response case
design and added additional values in version 15.0.1.

Added items 9-12 in version 23.4.0

Changed description of item 9 to “not used” in version
24.2.0

## See Also

[SetPreference](SetPreference_%7BConcrete_Indian_IS_456_2000%7D.md)
