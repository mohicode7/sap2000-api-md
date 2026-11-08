# SetPreference

## Syntax

SapObject.SapModel.DesignColdFormed.AISI\_ASD96.SetPreference

## VB6 Procedure

Function SetPreference(ByVal Item As Long, ByVal Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 10, inclusive, indicating the preference item considered.

1 = Framing type

2 = Demand/capacity ratio limit

3 = Omega bending stiffened

4 = Omega bending unstiffened

5 = Omega bending lateral torsional buckling

6 = Omega shear slender

7 = Omega shear nonslender

8 = Omega axial tension

9 = Omega axial compression

10 = Time history design

Value

The value of the considered preference item.

1 = Framing type

1 = Moment Frame

2 = Braced Frame

2 = Demand/capacity ratio limit

Value > 0

3 = Omega bending stiffened

Value > 0

4 = Omega bending unstiffened

Value > 0

5 = Omega bending lateral torsional buckling

Value > 0

6 = Omega shear slender

Value > 0

7 = Omega shear nonslender

Value > 0

8 = Omega axial tension

Value > 0

9 = Omega axial compression

Value > 0

10 = Time history design

1 = Envelopes

2 = Step-by step

## Remarks

This function sets the value of a cold formed design preference item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetColdFormedDesignPreferenceItemAISI_ASD96()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'add cold formed material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_COLDFORMED, , , , MATERIAL_COLDFORMED_SUBTYPE_ASTM_A653SQGr50)



   'create new cold formed frame section property

      ret = SapModel.PropFrame.SetColdC("CdC", Name , 9, 3, 0.06, 0.25, 0.5)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 72, 2, 144, True, "CdC", "CdC")



   'set coldformed design code

      ret = SapModel.DesignColdFormed.SetCode("AISI-ASD96")



   'set preference item

      ret = SapModel.DesignColdFormed.AISI_ASD96.SetPreference(1, 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetPreference](GetPreference_%7BCold_Formed_AISI_ASD96%7D.md)
