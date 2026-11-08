# SetPreference

## Syntax

SapObject.SapModel.DesignSteel.SP\_16\_13330\_2011.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 15, inclusive, indicating the preference item considered.

1 = Multi-response case design

2 = Framing type

3 = Section class

4 = GammaM

5 = GammaC

6 = GammaU

7 = GammaC1

8 = Consider deflection

9 = DL deflection limit, L/Value

10 = SDL + LL deflection limit, L/Value

11 = LL deflection limit, L/Value

12 = Total load deflection limit, L/Value

13 = Total camber limit, L/Value

14 = Pattern live load factor

15 = Demand/capacity ratio limit

Value

The value of the considered preference item.

1 = Multi-response case design

   1 = Envelopes

   2 = Step-by-step

   3 = Last step

   4 = Envelopes -- All

   5 = Step-by-step -- All

2 = Framing type

   1 = Moment frame

   2 = Braced frame

3 = Section class

   1 = Class 1

   2 = Class 2

   3 = Class 3

   4 = Class 4

4 = GammaM

Value > 0

5 = GammaC

Value > 0

6 = GammaU

Value > 0

7 = GammaC1

      Value > 0

8 = Consider deflection

0 = No

Any other value = Yes

9 = DL deflection limit, L/Value

Value > 0

10 = SDL + LL deflection limit, L/Value

  Value > 0

11 = LL deflection limit, L/Value

  Value > 0

12 = Total deflection limit, L/Value

  Value > 0

13 = Total camber limit, L/Value

  Value > 0

14 = Pattern live load factor

  Value >= 0

15 = Demand/capacity ratio limit

  Value > 0

## Remarks

This function retrieves the value of a steel design preference item.

The function returns zero if the item is successfully retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetSteelDesignPreferenceItemSP_16_13330_2011()

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

      ret = SapModel.DesignSteel.SetCode("SP 16.13330.2011")



   'set preference item

      ret = SapModel.DesignSteel.SP_16_13330_2011.SetPreference(4, 1.1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 20.1.0.

## See Also

[GetPreference](GetPreference_%7BSP_16-13330-2011%7D.md)
