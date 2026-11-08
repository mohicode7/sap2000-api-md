# SetOverwrite

## Syntax

SapObject.SapModel.DesignConcreteShell.Eurocode\_2\_2004.SetOverwrite

## VB6 Procedure

Function SetOverwrite(ByVal Name As String, ByVal Item As Long, ByVal Value As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object with a concrete shell design procedure.

Item

This is an integer between 1 and 5, inclusive, indicating the overwrite item considered.

1 = Number rebar layer

2 = Concrete cover to center of top rebar layer in direction 1

3 = Concrete cover to center of top rebar layer in direction 2

4 = Concrete cover to center of bottom rebar layer in direction 1

5 = Concrete cover to center of bottom rebar layer in direction 2

Value

The value of the considered overwrite item.

1 = Number rebar layer

0 = Program Default

1 = 1 layer

2 = 2 layers

2 = Concrete cover to center of top rebar layer in direction 1

Value >= 0; 0 means value taken from shell section definition.

3 = Concrete cover to center of top rebar layer in direction 2

Value >= 0; 0 means value taken from shell section definition.

4 = Concrete cover to center of bottom rebar layer in direction 1

Value >= 0; 0 means value taken from shell section definition.

5 = Concrete cover to center of bottom rebar layer in direction 2

Value >= 0; 0 value taken from shell section definition

This function sets the value of a concrete design overwrite item.

The function returns zero if the item is successfully set; otherwise it returns a nonzero value.

## VBA Example







```
Sub SetConcreteShellDesignOverwriteItemEurocode_2_2004()

    'dimension variables

    Dim SapObject as cOAPI

    Dim SapModel As cSapModel

    Dim ret As Long
      Dim value As Double
      Dim progdet As Boolean



    'create Sap2000 object

    Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



    'start Sap2000 application

    SapObject.ApplicationStart



    'create SapModel object

    Set SapModel = SapObject.SapModel
  
'initialize model
      ret = SapModel.InitializeNewModel(eUnits.kip_in_F)
   'create a wall model from template
      ret = SapModel.File.NewWall(6, 4, 6, 4)
    'set concrete shell design code

    ret = SapModel.DesignConcreteShell.SetCode("Eurocode 2 2004")

   'set concrete shell design overwrite
      ret = SapModel.DesignConcreteShell.Eurocode_2_2004.SetOverwrite("1", 4, 1.345)



    'get overwrite item

    ret = SapModel.DesignConcreteShell.Eurocode_2_2004.GetOverwrite("1", 4, value, progdet)



    'close Sap2000

    SapObject.ApplicationExit(False)

    Set SapModel = Nothing

    Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 26.2.0.

## See Also

[GetOverwrite](GetOverwrite_%7BConcrete_Shell_Design_Eurocode_2-2004%7D.md)
