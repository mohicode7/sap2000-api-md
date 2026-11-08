# NewBeam

## Syntax

SapObject.SapModel.File.NewBeam

## VB6 Procedure

Function NewBeam(ByVal NumberSpans As Long, ByVal SpanLength As Double, Restraint As Boolean = True, Optional ByVal Beam As String = "Default") As Long

## Parameters

NumberSpans

The number of spans for the beam.

SpanLength

The length of each span. [L]

Restraint

Joint restraints are provided at the ends of each span when this item is True.

Beam

The frame section property used for the beam. This must either be Default or the name of a defined frame section property.

## Remarks

Do not use this function to add to an existing model. This function should be used only for creating a new model and typically would be preceded by calls to ApplicationStart or InitializeNewModel.

The function returns zero if the new beam model is successfully created, otherwise it returns a nonzero value.

## VBA Example



```
Sub NewBeam()

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



   'create a beam from template

      ret = SapModel.File.NewBeam(3, 30)



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

[New2DFrame](New2DFrame.md)

[New3DFrame](New3DFrame.md)

[NewSolidBlock](NewSolidBlock.md)

[NewWall](NewWall.md)

[ApplicationStart](../General_Functions/ApplicationStart.md)

[InitializeNewModel](../General_Functions/InitializeNewModel.md)
