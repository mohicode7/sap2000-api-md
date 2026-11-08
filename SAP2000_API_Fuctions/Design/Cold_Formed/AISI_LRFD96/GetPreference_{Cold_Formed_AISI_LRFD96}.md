# GetPreference

## Syntax

SapObject.SapModel.DesignColdFormed.AISI\_LRFD96.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 10, inclusive, indicating the preference item considered.

1 = Framing type

2 = Demand/capacity ratio limit

3 = Phi bending stiffened

4 = Phi bending unstiffened

5 = Phi bending lateral torsional buckling

6 = Phi shear slender

7 = Phi shear nonslender

8 = Phi axial tension

9 = Phi axial compression

10 = Time history design

Value

The value of the considered preference item.

1 = Framing type

1 = Moment Frame

2 = Braced Frame

2 = Demand/capacity ratio limit

Value > 0

3 = Phi bending stiffened

Value > 0

4 = Phi bending unstiffened

Value > 0

5 = Phi bending lateral torsional buckling

Value > 0

6 = Phi shear slender

Value > 0

7 = Phi shear nonslender

Value > 0

8 = Phi axial tension

Value > 0

9 = Phi axial compression

Value > 0

10 = Time history design

1 = Envelopes

2 = Step-by step

## Remarks

This function retrieves the value of a cold formed design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetColdFormedDesignPreferenceItemAISI_LRFD96()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim Value As Double



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

      ret = SapModel.DesignColdFormed.SetCode("AISI-LRFD96")



   'get preference item

      ret = SapModel.DesignColdFormed.AISI_LRFD96.GetPreference(1, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetPreference](SetPreference_%7BCold_Formed_AISI_LRFD96%7D.md)
