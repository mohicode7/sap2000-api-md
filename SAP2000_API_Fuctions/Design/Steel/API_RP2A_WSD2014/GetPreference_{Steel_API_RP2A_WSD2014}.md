# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.API\_RP2A\_WSD2014.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 11, inclusive, indicating the preference item considered.

1 = Framing type

2 = Lateral factor, L/Value

3 = Consider deflection

4 = DL deflection limit, L/Value

5 = SDL + LL deflection limit, L/Value

6 = LL deflection limit, L/Value

7 = Total deflection limit, L/Value

8 = Total camber limit, L/Value

9 = Pattern live load factor

10 = Demand/capacity ratio limit

11 = Multi-response case design

Value

The value of the considered preference item.

1 = Framing type

1 = Moment Frame

2 = Braced Frame

2 = Lateral Factor

Value > 0

3 = Consider deflection

0 = No

Any other value = Yes

4 = DL deflection limit, L/Value

Value > 0

5 = SDL + LL deflection limit, L/Value

Value > 0

6 = LL deflection limit, L/Value

Value > 0

7 = Total deflection limit, L/Value

Value > 0

8 = Total camber limit, L/Value

Value > 0

9 = Pattern live load factor

Value >= 0

10 = Demand/capacity ratio limit

Value > 0

11 = Multi-response case design

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
Sub GetSteelDesignPreferenceItemAPI_RP2A_WSD2014()

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

      ret = SapModel.DesignSteel.SetCode("API RP2A-WSD2014")



   'get preference item

      ret = SapModel.DesignSteel.API_RP2A_WSD2014.GetPreference(1, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.1.0.

## See Also

[SetPreference](SetPreference_%7BSteel_API_RP2A_WSD2014%7D.md)
