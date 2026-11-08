# Trim

## Syntax

SapObject.SapModel.EditFrame.Trim

## VB6 Procedure

Function Trim(ByVal Name As String, ByVal IEnd As Boolean, ByVal JEnd As Boolean, ByVal Item1 As String, Optional ByVal Item2 As String = "") As Long

## Parameters

Name

The name of an existing straight frame object to be trimmed.

IEnd

This item is True if the I-End of the frame object specified by the Name item is to be trimmed.

JEnd

This item is True if the J-End of the frame object specified by the Name item is to be trimmed.

Item1

The name of an existing straight frame object used as a trim line.

Item2

The name of an existing straight frame object used as a trim line.

## Remarks

This function trims straight frame objects. Curved frame objects are not trimmed.

The function returns zero if the frame objects are successfully trimmed; otherwise it returns a nonzero value.

## VBA Example



```
Sub TrimFrameObject()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



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



   'add frame object by coordinates

      ret = SapModel.FrameObj.AddByCoord(-180, 0, 100, -180, 0, 360, Name)



   'refresh window

      ret = SapModel.View.RefreshWindow



   'trim frame object

      ret = SapModel.EditFrame.Trim(Name, True, True, "7", "8")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[Extend](Extend.md)
