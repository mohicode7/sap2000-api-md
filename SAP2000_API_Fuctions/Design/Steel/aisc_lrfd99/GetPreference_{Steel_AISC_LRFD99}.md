# GetPreference

## Syntax

SapObject.SapModel.DesignSteel.AISC\_LRFD99.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 21, inclusive, indicating the preference item considered.

1 = Framing type

2 = Seismic design capacity

3 = Phi bending

4 = Phi compression

5 = Phi tension - yielding

6 = Phi tension - fracture

7 = Phi shear

8 = Phi shear - torsion

9 = Phi compression, angle

10 = Ignore seismic code

11 = Ignore special seismic code

12 = Is doubler plate plug-welded

13 = Consider deflection

14 = DL deflection limit, L/Value

15 = SDL + LL deflection limit, L/Value

16 = LL deflection limit, L/Value

17 = Total deflection limit, L/Value

18 = Total camber limit, L/Value

19 = Pattern live load factor

20 = Demand capacity ratio limit

21 = Multi-response case design

Value

The value of the considered preference item.

1 = Framing type

1 = OMF

2 = IMF

3 = SMF

4 = OCBF

5 = SCBF

6 = EBF

2 = Seismic design capacity

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

3 = Phi bending

 Value > 0

4 = Phi compression

Value > 0

5 = Phi tension – yielding

Value > 0

6 = Phi tension - fracture

Value > 0

7 = Phi shear

Value > 0

8 = Phi shear - torsion

Value > 0

9 = Phi compression, angle

Value > 0

10 = Ignore seismic code

0 = No

Any other value = Yes

11 = Ignore special seismic code

0 = No

Any other value = Yes

12 = Is doubler plate plug-welded

0 = No

Any other value = Yes

13 = Consider deflection

0 = No

Any other value = Yes

14 = DL deflection limit, L/Value

Value > 0

15 = SDL + LL deflection limit, L/Value

Value > 0

16 = LL deflection limit, L/Value

Value > 0

17 = Total deflection limit, L/Value

Value > 0

18 = Total camber limit, L/Value

Value > 0

19 = Pattern live load factor

Value >= 0

20 = Demand capacity ratio limit

Value > 0

21 = Multi-response case design

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
Sub GetSteelDesignPreferenceItemAISC_LRFD99()

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

      ret = SapModel.DesignSteel.SetCode("AISC-LRFD99")



   'get preference item

      ret = SapModel.DesignSteel.AISC_LRFD99.GetPreference(1, Value)



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

[SetPreference](SetPreference_%7BSteel_AISC_LRFD99%7D.md)
