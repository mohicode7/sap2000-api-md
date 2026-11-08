# SetTargetDispl

## Syntax

SapObject.SapModel.DesignSteel.SetTargetDispl

## VB6 Procedure

Function SetTargetDispl(ByVal NumberItems As Long, ByRef LoadCase() As String, ByRef Point() As String, ByRef Displ() As Double, Optional ByVal Active As Boolean = True) As Long

## Parameters

NumberItems

The number of lateral displacement targets specified.

LoadCase

This is an array that includes the name of the static linear load case associated with each lateral displacement target.

Point

This is an array that includes the name of the point object associated to which the lateral displacement target applies.

Displ

This is an array that includes the lateral displacement target. [L]

Active

If this item is True, all specified lateral displacement targets are active. If it is False, they are inactive.

## Remarks

This function sets lateral displacement targets for steel design.

The function returns zero if the targets are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignTargetDispl()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyLCase() As String

      Dim MyPoint() As String

      Dim MyDispl() As Double



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



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("EQX", LTYPE_QUAKE)



   'assign UBC97 parameters

      ret = SapModel.LoadPatterns.AutoSeismic.SetUBC97("EQX", 1, 0.05, 1, 0.035, 0, False, 0, 0, 1, 3, 0.4, 0, 0, 1, 3, 5, 0, 0, 1.15, 6)



   'set target displacement data

      ReDim MyLCase(0)

      ReDim MyPoint(0)

      ReDim MyDispl(0)

      MyLCase(0) = "EQX"

      MyPoint(0) = "3"

      MyDispl(0) = 0.4

      ret = SapModel.DesignSteel.SetTargetDispl(1, MyLCase, MyPoint, MyDispl)



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

[GetTargetDispl](GetTargetDispl.md)
