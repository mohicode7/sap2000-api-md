# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.Chinese\_2000.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 14, inclusive, indicating the preference item considered.

1 = Framing type

2 = Gamma0

3 = Ignore b/t check

4 = Classify beam as flexo compression member

5 = Consider deflection

6 = DL deflection limit, L/Value

7 = SDL + LL deflection limit, L/Value

8 = LL deflection limit, L/Value

9 = Total load deflection limit, L/Value

10 = Total camber limit, L/Value

11 = Pattern live load factor

12 = Demand/capacity ratio limit

13 = Multi-response case design

14 = Is tall building?

Value

The value of the considered preference item.

1 = Framing type

0 = As specified in preferences

1 = Sway Moment Frame, SMF

2 = Concentrically Braced Frame, CBF

3 = Eccentrically Braced Frame, EBF

4 = NonSway Moment Frame, NMF

2 = Gamma0

Value > 0

3 = Ignore b/t check

0 = No

Any other value = Yes

4 = Classify beam as flexo compression member

0 = No

Any other value = Yes

5 = Consider deflection

0 = No

Any other value = Yes

6 = DL deflection limit, L/Value

Value > 0

7 = SDL + LL deflection limit, L/Value

Value > 0

8 = LL deflection limit, L/Value

Value > 0

9 = Total load deflection limit, L/Value

Value > 0

10 = Total camber limit, L/Value

Value > 0

11 = Pattern live load factor

Value >= 0

12 = Demand/capacity ratio limit

Value > 0

13 = Multi-response case design

  1 = Envelopes

  2 = Step-by-step

  3 = Last step

  4 = Envelopes -- All

  5 = Step-by-step -- All

14 = Tall building

        0 = No

        1 = Yes

## Remarks

This function retrieves the value of a steel design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItemChinese_2002()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Value As Double



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'createSapModelobject

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set steel design code

      ret = SapModel.DesignSteel.SetCode("Chinese 2002")



   'get preference item

      ret = SapModel.DesignSteel.Chinese_2002.GetPreference(1, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Added item 14 in Version 14.0.0.

Modified Item 1 in version 14.1.0.

Changed Time history design item to Multi-response case design and added additional values in version 15.0.1.

## See Also

[SetPreference](SetPreference_%7BSteel_Chinese_2002%7D.md)
