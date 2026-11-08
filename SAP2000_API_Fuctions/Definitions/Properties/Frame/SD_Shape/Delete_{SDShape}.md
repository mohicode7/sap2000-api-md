# Delete

## Syntax

SapObject.SapModel.PropFrame.SDShape.Delete

## VB6 Procedure

Function Delete(ByVal Name As String, ByRef ShapeName As String, Optional ByVal All As Boolean = False) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing shape in a section designer property. If the All item is True, this item may be specified as a blank string.

All

If this item is True, all shapes in the section designer property specified by the Name item are deleted.

## Remarks

This function deletes shapes from a section designer property.

The function returns zero if the shape is successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteFrameSDPropShape()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", "A992Fy50")



   'add I-section shapes to new property

      ret = SapModel.PropFrame.SDShape.SetISection("SD1", "SH1", "A992Fy50", "", 0, -9, 0, -1, 18, 6, 1, 0.5, 6, 1)

      ret = SapModel.PropFrame.SDShape.SetISection("SD1", "SH2", "A992Fy50", "", -9, -9, 0, -1, 18, 6, 1, 0.5, 6, 1)

      ret = SapModel.PropFrame.SDShape.SetISection("SD1", "SH3", "A992Fy50", "", 9, -9, 0, -1, 18, 6, 1, 0.5, 6, 1)



   'delete shape from section designer property

      ret = SapModel.PropFrame.SDShape.Delete("SD1", "SH1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

Modified optional argument All to be ByVal in version 12.0.1.

## See Also
