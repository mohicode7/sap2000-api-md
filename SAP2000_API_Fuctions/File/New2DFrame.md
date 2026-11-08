# New2DFrame

## Syntax

SapObject.SapModel.File.New2DFrame

## VB6 Procedure

Function New2DFrame(ByVal TempType As e2DFrameType, ByVal NumberStorys As Long, ByVal StoryHeight As Double, ByVal NumberBays As Long, ByVal BayWidth As Double, Optional ByVal Restraint As Boolean = True, Optional ByVal Beam As String = "Default", Optional ByVal Column As String = "Default", Optional ByVal Brace As String = "Default") As Long

## Parameters

TempType

One of the following 2D frame template types in the e2DFrameType enumeration.

PortalFrame = 0

ConcentricBraced = 1

EccentricBraced = 2

NumberStorys

The number of stories in the frame.

StoryHeight

The height of each story. [L]

NumberBays

The number of bays in the frame.

BayWidth

The width of each bay. [L]

Restraint

Joint restraints are provided at the base of the frame when this item is True.

Beam

The frame section property used for all beams in the frame. This must either be Default or the name of a defined frame section property.

Column

The frame section property used for all columns in the frame. This must either be Default or the name of a defined frame section property.

Brace

The frame section property used for all braces in the frame. This must either be Default or the name of a defined frame section property. This item does not apply to the portal frame.

## Remarks

Do not use this function to add to an existing model. This function should be used only for creating a new model and typically would be preceded by calls to ApplicationStart or InitializeNewModel.

The function returns zero if the new 2D frame model is successfully created, otherwise it returns a nonzero value.

## VBA Example



```
Sub New2DFrameModel()

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



   'create a 2D frame model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 12, 3, 28)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[New3DFrame](New3DFrame.md)

[NewBeam](NewBeam.md)

[NewBlank](NewBlank.md)

[NewSolidBlock](NewSolidBlock.md)

[NewWall](NewWall.md)

[ApplicationStart](../General_Functions/ApplicationStart.md)

[InitializeNewModel](../General_Functions/InitializeNewModel.md)
