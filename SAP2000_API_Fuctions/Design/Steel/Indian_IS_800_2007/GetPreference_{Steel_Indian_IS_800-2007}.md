# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.Indian\_IS\_800\_2007.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 18, inclusive, indicating the preference item considered.

1 = Framing type

2 = Importance factor

3 = Seismic Zone

4 = Consider P-delta Done

5 = GammaM0

6 = GammaM1

7 = Ignore sseismic code

8 = Ignore special seismic load

9 = Is doubler plate plug-welded

10 = Consider deflection

11 = DL deflection limit, L/Value

12 = SDL + LL deflection limit, L/Value

13 = LL deflection limit, L/Value

14 = Total deflection limit, L/Value

15 = Total camber limit, L/Value

16 = Pattern live load factor

17 = Demand/capacity ratio limit

18 = Multi-response case design

Value

The value of the considered preference item.

1 = Framing type

1 = SMF

2 = OMF

3 = SCBF

4 = OCBF

5 = EBF

6 = Secondary

2 = Importance factor

Value > 0

3 = Seismic Zone

1 = Zone I

2 = Zone II

3 = Zone III

4 = Zone IV

5 = Zone V

4 = Consider P-delta Done

0 = No

Any other value = Yes

5 = GammaM0

Value > 0

6 = GammaM1

Value > 0

7 = Ignore sseismic code

0 = No

Any other value = Yes

8 = Ignore special seismic load

0 = No

Any other value = Yes

9 = Is doubler plate plug-welded

0 = No

Any other value = Yes

10 = Consider deflection

 0 = No

 Any other value = Yes

11 = DL deflection limit, L/Value

  Value > 0

12 = SDL + LL deflection limit, L/Value

  Value > 0

13 = LL deflection limit, L/Value

  Value > 0

14 = Total deflection limit, L/Value

  Value > 0

15 = Total camber limit, L/Value

  Value > 0

16 = Pattern live load factor

  Value >= 0

17 = Demand/capacity ratio limit

  Value > 0

18 = Time history design

  1 = Envelopes

  2 = Step-by step

  3 = Last step

  4 = Envelopes - All

  5 = Step-by step - All

## Remarks

This function retrieves the value of a steel design preference item.

The function returns zero if the item is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItemIndian_IS_800_2007()

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

      ret = SapModel.DesignSteel.SetCode("Indian IS 800:2007")



   'get preference item

      ret = SapModel.DesignSteel.Indian_IS_800_2007.GetPreference(1, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.1.

## See Also

[SetPreference](SetPreference_%7BSteel_Indian_IS_800-2007%7D.md)
