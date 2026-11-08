# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.CISC\_95.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 14, inclusive, indicating the preference item considered.

1 = Framing type

2 = Phi bending

3 = Phi compression

4 = Phi tension

5 = Phi shear

6 = Consider deflection

7 = DL deflection limit, L/Value

8 = SDL + LL deflection limit, L/Value

9 = LL deflection limit, L/Value

10 = Total deflection limit, L/Value

11 = Total camber limit, L/Value

12 = Pattern live load factor

13 = Demand/capacity ratio limit

14 = Multi-response case design

Value

The value of the considered preference item.

1 = Framing type

1 = Moment Frame

2 = Braced Frame

2 = Phi bending

Value > 0

3 = Phi compression

Value > 0

4 = Phi tension

Value > 0

5 = Phi shear

Value > 0

6 = Consider deflection

0 = No

Any other value = Yes

7 = DL deflection limit, L/Value

Value > 0

8 = SDL + LL deflection limit, L/Value

Value > 0

9 = LL deflection limit, L/Value

Value > 0

10 = Total load deflection limit, L/Value

Value > 0

11 = Total camber limit, L/Value

Value > 0

12 = Pattern live load factor

Value >= 0

13 = Demand/capacity ratio limit

Value > 0

14 = Multi-response case design

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
Sub GetSteelDesignPreferenceItemCISC_95()

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

      ret = SapModel.DesignSteel.SetCode("CISC 95")



   'get preference item

      ret = SapModel.DesignSteel.CISC_95.GetPreference(1, Value)



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

[SetPreference](SetPreference_%7BSteel_CISC_95%7D.md)
