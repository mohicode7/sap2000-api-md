# SetPreference

## Syntax

SapObject.SapModel.DesignConcrete.NZS\_3101\_95.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 14, inclusive, indicating the preference item considered.

1 = Number of interaction curves

2 = Number of interaction points

3 = Consider minimum eccentricity

4 = Phi bending

5 = Phi tension

6 = Phi compression

7 = Phi shear

8 = Omega

9 = Phi 0

10 = Rm

11 = Rv

12 = Pattern live load factor

13 = Utilization factor limit

14 = Multi-response case design

Value

The value of the considered preference item.

1 = Number of interaction curves

Value >= 4 and devisable by 4

2 = Number of interaction points

Value >= 5 and odd

3 = Consider minimum eccentricity

0 = No

Any other value = Yes

4 = Phi bending

Value > 0

5 = Phi tension

Value > 0

6 = Phi compression

Value > 0

7 = Phi shear

Value > 0

8 = Omega

Value > 0

9 = Phi 0

Value > 0

10 = Rm

Value > 0

11 = Rv

Value > 0

12 = Pattern live load factor

Value >= 0

13 = Utilization factor limit

Value > 0

14 = Multi-response case design

  1 = Envelopes

  2 = Step-by-step

  3 = Last step

  4 = Envelopes -- All

  5 = Step-by-step -- All

## Remarks

This function sets the value of a concrete design preference item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignPreferenceItemNZS_3101_95()

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

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("NZS 3101-95")



   'set preference item

      ret = SapModel.DesignConcrete.NZS_3101_95.SetPreference(2, 9)



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

[GetPreference](GetPreference_%7BConcrete_NZS_3101_95%7D.md)
