# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.UBC97\_LRFD.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 12, inclusive, indicating the preference item considered.

1 = Framing type

2 = Importance factor

3 = Seismic zone

4 = Phi bending

5 = Phi compression

6 = Phi tension

7 = Phi shear

8 = Phi compression angle

9 = Consider deflection

10 = DL deflection limit, L/Value

11 = SDL + LL deflection limit, L/Value

12 = LL deflection limit, L/Value

13 = Total deflection limit, L/Value

14 = Total camber limit, L/Value

15 = Pattern live load factor

16 = Demand/capacity ratio limit

17 = Multi-response case design

Value

The value of the considered preference item.

1 = Framing type

1 = Ordinary MRF

2 = Special MRF

3 = Braced Frame

4 = Special CBF

5 = EBF

2 = Importance factor

Value > 0

3 = Seismic zone

1 = Zone 0

2 = Zone 1

3 = Zone 2

4 = Zone 3

5 = Zone 4

3 = Lateral factor

Value > 0

4 = Phi bending

   Value > 0

5 = Phi compression

Value > 0

6 = Phi tension

Value > 0

7 = Phi shear

Value > 0

8 = Phi compression angle

Value > 0

9 = Consider deflection

0 = No

Any other value = Yes

10 = DL deflection limit, L/Value

Value > 0

11 = SDL + LL deflection limit, L/Value

Value > 0

12 = LL deflection limit, L/Value

Value > 0

13 = Total deflection limit, L/Value

Value > 0

14 = Total camber limit, L/Value

Value > 0

15 = Pattern live load factor

Value >= 0

16 = Demand/capacity ratio limit

Value > 0

17 = Multi-response case design

  1 = Envelopes

  2 = Step-by-step

  3 = Last step

  4 = Envelopes -- All

  5 = Step-by-step -- All

## Remarks

This function retrieves the value of a steel design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItemUBC97_LRFD()

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("UBC97-LRFD")



   'get preference item

      ret = SapModel.DesignSteel.UBC97_LRFD.GetPreference(1, Value)



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

[SetPreference](SetPreference_%7BSteel_UBC97-LRFD%7D.md)
