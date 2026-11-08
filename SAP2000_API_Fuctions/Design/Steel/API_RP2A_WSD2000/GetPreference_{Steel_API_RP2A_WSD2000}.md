# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.API\_RP2A\_WSD2000.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 12, inclusive, indicating the preference item considered.

1 = Framing type

2 = Tubular joint punching load method

3 = Lateral factor, L/Value

4 = Consider deflection

5 = DL deflection limit, L/Value

6 = SDL + LL deflection limit, L/Value

7 = LL deflection limit, L/Value

8 = Total deflection limit, L/Value

9 = Total camber limit, L/Value

10 = Pattern live load factor

11 = Demand/capacity ratio limit

12 = Multi-response case design

13 = Code supplements

Value

The value of the considered preference item.

1 = Framing type

1 = Moment Frame

2 = Braced Frame

2 = Tubular joint punching load method

1 = Punching Shear

2 = Nominal Load

3 = Lateral Factor

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

13 = Code supplements

1 = None

2 = Supplements 2 and 3

## Remarks

This function retrieves the value of a steel design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItemAPI_RP2A_WSD2000()

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

      ret = SapModel.DesignSteel.SetCode("API RP2A-WSD2000")



   'get preference item

      ret = SapModel.DesignSteel.API_RP2A_WSD2000.GetPreference(1, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed Time history design item to Multi-response case design and added additional values in version 15.0.1.

Added code supplements item in version 16.1.0.

## See Also

[SetPreference](SetPreference_%7BStl_API_RP2A_WSD2000%7D.md)
