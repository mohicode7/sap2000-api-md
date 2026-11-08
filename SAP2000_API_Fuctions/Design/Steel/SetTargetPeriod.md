# SetTargetPeriod

## Syntax

SapObject.SapModel.DesignSteel.SetTargetPeriod

## VB6 Procedure

Function SetTargetPeriod(ByVal NumberItems As Long, ByVal ModalCase As String, ByRef Mode() As Long, ByRef Period() As Double, Optional ByVal Active As Boolean = True) As Long

## Parameters

NumberItems

The number of lateral displacement targets specified.

ModalCase

The name of the modal load case for which the target periods apply.

Mode

This is an array that includes the mode number associated with each target period.

Period

This is an array that includes the target periods. [s]

Active

If this item is True, all specified target periods are active. If it is False, they are inactive.

## Remarks

This function sets time period targets for steel design.

The function returns zero if the targets are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignTargetPeriod()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyMode() As Long

      Dim MyPeriod() As Double



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



   'set target period data

      ReDim MyLCase(0)

      ReDim MyMode(1)

      ReDim MyPeriod(1)

      MyMode(0) = 1

      MyPeriod(0) = 0.6

      MyMode(1) = 2

      MyPeriod(1) = 0.5

      ret = SapModel.DesignSteel.SetTargetPeriod(2, "MODAL", MyMode, MyPeriod)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetTargetPeriod](GetTargetPeriod.md)
