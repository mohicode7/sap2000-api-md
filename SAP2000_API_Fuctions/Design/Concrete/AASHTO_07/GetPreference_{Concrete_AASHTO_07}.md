# GetPreference

## Syntax

SapObject.SapModel.DesignConcrete.AASHTO\_Concrete\_07.GetPreference

## VB6 Procedure

Function GetPreference(ByVal Item As Long, ByRef Value As Double) As Long

## Parameters

Item

This is an integer between 1 and 7, inclusive, indicating the preference item considered.

1 = Number of interaction curves

2 = Number of interaction points

3 = Consider minimum eccentricity

4 = Seismic Zone

5 = Pattern live load factor

6 = Utilization factor limit

7 = Multi-response case design

Value

The value of the considered preference item.

1 = Number of interaction curves

Value >= 4 and divisible by 4

2 = Number of interaction points

Value >= 5 and odd

3 = Consider minimum eccentricity

0 = No

Any other value = Yes

4 = Seismic Zone

  0 = Zone 0

  1 = Zone 1

  2 = Zone 2

  3 = Zone 3

  4 = Zone 4

5 = Pattern live load factor

Value >= 0

6 = Utilization factor limit

Value > 0

7 = Multi-response case design

1 = Envelopes

2 = Step-by-step

3 = Last step

4 = Envelopes -- All

5 = Step-by-step -- All

## Remarks

This function retrieves the value of a concrete design preference item.

The function returns zero if the item is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetConcreteDesignPreferenceItem AASHTO_Concrete_07()

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



   'create new concrete frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")



   'set concrete design code

      ret = SapModel.DesignConcrete.SetCode("AASHTO Concrete 07")



   'get preference item

      ret = SapModel.DesignConcrete.AASHTO_Concrete_07.GetPreference(2, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.1.0.

Changed Time history design item to Multi-response case design and added additional values in version 15.0.1.

## See Also

[SetPreference](SetPreference_%7BConcrete_AASHTO_07%7D.md)
