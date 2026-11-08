# GetPreference

## Syntax

SapObject.SapModel.DesignConcrete.ACI\_318\_02.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 13, inclusive, indicating the preference item considered.

1 = Number of interaction curves

2 = Number of interaction points

3 = Consider minimum eccentricity

4 = Seismic design category

5 = Phi tension controlled

6 = Phi compression controlled tied

7 = Phi compression controlled spiral

8 = Phi shear and/or torsion

9 = Phi shear seismic

10 = Phi joint shear

11 = Pattern live load factor

12 = Utilization factor limit

13 = Multi-response case design

Value

The value of the considered preference item.

1 = Number of interaction curves

Value >= 4 and devisable by 4

2 = Number of interaction points

Value >= 5 and odd

3 = Consider minimum eccentricity

0 = No

Any other value = Yes

4 = Seismic design category

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

5 = Phi tension controlled

Value > 0

6 = Phi compression controlled tied

Value > 0

7 = Phi compression controlled spiral

Value > 0

8 = Phi shear and/or torsion

Value > 0

9 = Phi shear seismic

Value > 0

10 = Phi joint shear

Value > 0

11 = Pattern live load factor

Value >= 0

12 = Utilization factor limit

Value > 0

13 = Multi-response case design

  1 = Envelopes

  2 = Step-by-step

  3 = Last step

  4 = Envelopes -- All

  5 = Step-by-step -- All

## Remarks

This function retrieves the value of a concrete design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignPreferenceItemACI_318_02()

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

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("ACI 318-02")



   'get preference item

      ret = SapModel.DesignConcrete.ACI_318_02.GetPreference(2, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed Time history design item to Multi-response case design and added additional values in version 15.0.1.

## See Also

[SetPreference](SetPreference_%7BConcrete_ACI_318-02%7D.md)
