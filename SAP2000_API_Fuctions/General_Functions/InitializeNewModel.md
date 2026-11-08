# InitializeNewModel

## Syntax

SapObject.SapModel.InitializeNewModel

## VB6 Procedure

Function InitializeNewModel(Optional ByVal Units As eUnits = kip\_in\_F) As Long

## Parameters

Units

This is the database units for the new model. All data is internally stored in the model in these units. The units are one of the following items in the eUnits enumeration:

lb\_in\_F = 1

lb\_ft\_F = 2

kip\_in\_F = 3

kip\_ft\_F = 4

kN\_mm\_C = 5

kN\_m\_C = 6

kgf\_mm\_C = 7

kgf\_m\_C = 8

N\_mm\_C = 9

N\_m\_C = 10

Ton\_mm\_C = 11

Ton\_m\_C = 12

kN\_cm\_C = 13

kgf\_cm\_C = 14

N\_cm\_C = 15

Ton\_cm\_C = 16

## Remarks

This function clears the previous model and initializes the program for a new model. If it is later needed, you should save your previous model prior to calling this function.

After calling the InitializeNewModel function, it is not necessary to also call the ApplicationStart function because the functionality of the ApplicationStart function is included in the InitializeNewModel function.

The function returns zero if a new model is successfully initialized, otherwise it returns a nonzero value.

## VBA Example



```
Sub InitializeNewModel()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel(kip_ft_F)



   'create model from template

      ret = SapModel.File.New2DFrame(ConcentricBraced, 3, 12, 3, 30)



   'save model

      ret = SapModel.File.Save("C:\SapAPI\MyFirstSapModel.sdb")



   'initialize new model

      SapModel.InitializeNewModel(kN_m_C)



   'create model from template

      ret = SapModel.File.New3DFrame(FlatPlate, 2, 4, 2, 10, 4, 10, False, , , , 2, 2)



   'save model

      ret = SapModel.File.Save("C:\SapAPI\MySecondSapModel.sdb")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[ApplicationStart](ApplicationStart.md)
