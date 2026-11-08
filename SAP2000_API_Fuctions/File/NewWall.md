# NewWall

## Syntax

SapObject.SapModel.File.NewWall

## VB6 Procedure

Function NewWall(ByVal NumberXDivisions As Long, ByVal DivisionWidthX As Double, ByVal NumberZDivisions As Long, ByVal DivisionWidthZ As Double, Optional ByVal Restraint As Boolean = True, Optional ByVal Area As String = "Default") As Long

## Parameters

NumberXDivisions

The number of area objects in the global X direction of the wall.

DivisionWidthX

The width of each area object measured in the global X direction. [L]

NumberZDivisions

The number of area objects in the global Z direction of the wall.

DivisionWidthZ

The height of each area object measured in the global Z direction. [L]

Restraint

Joint restraints are provided at the base of the wall when this item is True.

Area

The shell section property used for the wall. This must either be Default or the name of a defined shell section property.

## Remarks

Do not use this function to add to an existing model. This function should be used only for creating a new model and typically would be preceded by calls to ApplicationStart or InitializeNewModel.

The function returns zero if the new wall model is successfully created, otherwise it returns a nonzero value.

## VBA Example



```
Sub NewWall()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel(kip_ft_F)



   'create a wall model from template

      ret = SapModel.File.NewWall(6, 4, 6, 4)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[NewBlank](NewBlank.md)

[NewBeam](NewBeam.md)

[New2DFrame](New2DFrame.md)

[New3DFrame](New3DFrame.md)

[NewSolidBlock](NewSolidBlock.md)

[ApplicationStart](../General_Functions/ApplicationStart.md)

[InitializeNewModel](../General_Functions/InitializeNewModel.md)
