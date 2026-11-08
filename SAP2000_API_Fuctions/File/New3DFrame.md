# New3DFrame

## Syntax

SapObject.SapModel.File.New3DFrame

## VB6 Procedure

Function New3DFrame(ByVal TempType As e3DFrameType, ByVal NumberStorys As Long, ByVal StoryHeight As Double, ByVal NumberBaysX As Long, ByVal BayWidthX As Double, ByVal NumberBaysY As Long, ByVal BayWidthY As Double, Optional ByVal Restraint As Boolean = True, Optional ByVal Beam As String = "Default", Optional ByVal Column As String = "Default", Optional ByVal Area As String = "Default", Optional ByVal NumberXDivisions As Long = 4, Optional ByVal NumberYDivisions As Long = 4) As Long

## Parameters

TempType

One of the following 3D frame template types in the e3DFrameType enumeration.

OpenFrame = 0

PerimeterFrame = 1

BeamSlab = 2

FlatPlate = 3

NumberStorys

The number of stories in the frame.

StoryHeight

The height of each story. [L]

NumberBaysX

The number of bays in the global X direction of the frame.

BayWidthX

The width of each bay in the global X direction of the frame. [L]

NumberBaysY

The number of bays in the global Y direction of the frame.

BayWidthY

The width of each bay in the global Y direction of the frame. [L]

Restraint

Joint restraints are provided at the base of the frame when this item is True.

Beam

The frame section property used for all beams in the frame. This must either be Default or the name of a defined frame section property.

Column

The frame section property used for all columns in the frame. This must either be Default or the name of a defined frame section property.

Area

The shell section property used for all floor slabs in the frame. This must either be Default or the name of a defined shell section property. This item does not apply to the open and perimeter frames.

NumberXDivisions

The number of divisions for each floor area object in the global X direction. This item does not apply to the open and perimeter frames.

NumberYDivisions

The number of divisions for each floor area object in the global Y direction. This item does not apply to the open and perimeter frames.

## Remarks

Do not use this function to add to an existing model. This function should be used only for creating a new model and typically would be preceded by calls to ApplicationStart or InitializeNewModel.

The function returns zero if the new 3D frame model is successfully created, otherwise it returns a nonzero value.

## VBA Example



```
Sub New3DFrameModel()

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



   'create a 3D frame model from template

      ret = SapModel.File.New3DFrame(BeamSlab, 3, 12, 3, 28, 2, 36)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[NewBeam](NewBeam.md)

[NewBlank](NewBlank.md)

[New2DFrame](New2DFrame.md)

[NewSolidBlock](NewSolidBlock.md)

[NewWall](NewWall.md)

[ApplicationStart](../General_Functions/ApplicationStart.md)

[InitializeNewModel](../General_Functions/InitializeNewModel.md)
