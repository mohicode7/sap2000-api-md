# SetPreference

## Syntax

SapObject.SapModel.DesignConcrete.AS\_3600\_09.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 13, inclusive, indicating the preference item considered.

1 = Number of interaction curves

2 = Number of interaction points

3 = Consider minimum eccentricity

4 = Phi tension controlled

5 = Phi compression controlled

6 = Phi shear and/or torsion

7 = Phi shear seismic

8 = Phi joint shear

9 = Pattern live load factor

10 = Utilization factor limit

11 = Time history design

12 = Consider torsion

13 = Design for B/C capacity ratio

Value

The value of the considered preference item.

1 = Number of interaction curves

Value >= 4 and divisible by 4

2 = Number of interaction points

Value >= 5 and odd

3 = Consider minimum eccentricity

0 = No

Any other value = Yes

4 = Phi tension controlled

Value > 0

5 = Phi compression controlled

Value > 0

6 = Phi shear and/or torsion

Value > 0

7 = Phi shear seismic

Value > 0

8 = Phi joint shear

Value > 0

9 = Pattern live load factor

Value >= 0

10 = Utilization factor limit

  Value > 0

11 = Multi-response case design

        1 = Envelopes

        2 = Step-by step

        3= Last step

        4 = Envelopes - All

        5 = Step-by step - All

12 = Consider torsion

0 = No

Any other value = Yes

13 = Design for B/C capacity ratio

0 = No

Any other value = Yes

## Remarks

This function sets the value of a concrete design preference item.

The function returns zero if the item is successfully set; otherwise, it returns a nonzero value.

## VBA Example



```
Sub SetConcreteDesignPreferenceItemAS_3600_09()

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



       'create new concrete frame section property

          ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



       'set concrete design code

          ret = SapModel.DesignConcrete.SetCode("AS 3600-00")



       'set preference item

          ret = SapModel.DesignConcrete.AS_3600_09.SetPreference(2, 9)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 15.01.

Added items 12~13 in version 23.4.0

## See Also

[GetPreference](mk:@MSITStore:W:/Released/SAP2000 V15/15.0.0 Server/CSi_OAPI_Documentation.chm::/SAP2000_API_Fuctions/Design/Concrete/Australian_AS_3600-01/GetPreference_%7bConcrete_Australian_AS_3600-01%7d.htm)
