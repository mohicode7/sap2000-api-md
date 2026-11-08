# SetPreference

## Syntax

SapObject.SapModel.DesignSteel.AISC\_LRFD93.SetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 15, inclusive, indicating the preference item considered.

1 = Framing type

2 = Phi bending

3 = Phi compression

4 = Phi tension

5 = Phi shear

6 = Phi compression, angle

7 = Consider deflection

8 = DL deflection limit, L/Value

9 = SDL + LL deflection limit, L/Value

10 = LL deflection limit, L/Value

11 = Total deflection limit, L/Value

12 = Total camber limit, L/Value

13 = Pattern live load factor

14 = Demand/capacity ratio limit

15 = Multi-response case design

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

6 = Phi compression, angle

Value > 0

7 = Consider deflection

0 = No

Any other value = Yes

8 = DL deflection limit, L/Value

Value > 0

9 = SDL + LL deflection limit, L/Value

Value > 0

10 = LL deflection limit, L/Value

  Value > 0

11 = Total deflection limit, L/Value

  Value > 0

12 = Total camber limit, L/Value

Value > 0

13 = Pattern live load factor

Value >= 0

14 = Demand/capacity ratio limit

Value > 0

15 = Multi-response case design

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
Sub SetSteelDesignPreferenceItemAISC_LRFD93 ()

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

      ret = SapModel.DesignSteel.SetCode("AISC-LRFD93")



   'set preference item

      ret = SapModel.DesignSteel.AISC_LRFD93.SetPreference(1, 2)



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

[GetPreference](GetPreference_%7BSteel_AISC_LRFD93%7D.md)
