# SetPreference

## Syntax

SapObject.SapModel.DesignSteel.EUROCODE\_3\_1993.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 12, inclusive, indicating the preference item considered.

1 = Framing type

2 = GammaM0

3 = GammeM1

4 = Consider deflection

5 = DL deflection limit, L/Value

6 = SDL + LL deflection limit, L/Value

7 = LL deflection limit, L/Value

8 = Total deflection limit, L/Value

9 = Total camber limit, L/Value

10 = Pattern live load factor

11 = Demand/capacity ratio limit

12 = Multi-response case design

Value

The value of the considered preference item.

1 = Framing type

1 = Moment Frame

2 = Braced Frame

2 = GammaM0

Value > 0

3 = GammeM1

Value > 0

4 = Consider deflection

0 = No

Any other value = Yes

5 = DL deflection limit, L/Value

Value > 0

6 = SDL + LL deflection limit, L/Value

Value > 0

7 = LL deflection limit, L/Value

Value > 0

8 = Total deflection limit, L/Value

Value > 0

9 = Total camber limit, L/Value

Value > 0

10 = Pattern live load factor

Value >= 0

11 = Demand/capacity ratio limit

Value > 0

12 = Multi-response case design

  1 = Envelopes

  2 = Step-by-step

  3 = Last step

  4 = Envelopes -- All

  5 = Step-by-step -- All

## Remarks

This function sets the value of a steel design preference item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignPreferenceItemEUROCODE_3_1993()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("EUROCODE 3-1993")



   'set preference item

      ret = SapModel.DesignSteel.EUROCODE_3_1993.SetPreference(1, 2)



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

[GetPreference](GetPreference_%7BSteel_Eurocode_3-1993%7D.md)
