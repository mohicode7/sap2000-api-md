# NewSolidBlock

## Syntax

SapObject.SapModel.File.NewSolidBlock

## VB6 Procedure

Function NewSolidBlock(ByVal XWidth As Double, ByVal YWidth As Double, ByVal Height As Double, Optional ByVal Restraint As Boolean = True, Optional ByVal Solid As String = "Default", Optional ByVal NumberXDivisions As Long = 5, Optional ByVal NumberYDivisions As Long = 8, Optional ByVal NumberZDivisions As Long = 10) As Long

## Parameters

XWidth

The total width of the solid block measured in the global X direction. [L]

YWidth

The total width of the solid block measured in the global Y direction. [L]

Height

The total height of the solid block measured in the global Z direction. [L]

Restraint

Joint restraints are provided at the base of the solid block when this item is True.

Solid

The solid property used for the solid block. This must either be Default or the name of a defined solid property.

NumberXDivisions

The number of solid objects in the global X direction of the block.

NumberYDivisions

The number of solid objects in the global Y direction of the block.

NumberZDivisions

The number of solid objects in the global Z direction of the block.

## Remarks

The function returns zero if the new solid block model is successfully created, otherwise it returns a nonzero value.

## VBA Example



```
Sub NewSolidBlock()

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



   'create a solid model from template

      ret = SapModel.File.NewSolidBlock(20, 50, 20)



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

[NewWall](NewWall.md)
