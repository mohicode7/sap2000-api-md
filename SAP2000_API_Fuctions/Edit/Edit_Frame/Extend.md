# Extend

## Syntax

SapObject.SapModel.EditFrame.Extend

## VB6 Procedure

Function Extend(ByVal Name As String, ByVal IEnd As Boolean, ByVal JEnd As Boolean, ByVal Item1 As String, Optional ByVal Item2 As String = "") As Long

## Parameters

Name

The name of an existing straight frame object to be extended.

IEnd

This item is True if the I-End of the frame object specified by the Name item is to be extended.

JEnd

This item is True if the J-End of the frame object specified by the Name item is to be extended.

Item1

The name of an existing straight frame object used as a extension line.

Item2

The name of an existing straight frame object used as a extension line.

## Remarks

This function extends straight frame objects. Curved frame objects are not extended.

The function returns zero if the frame objects are successfully extended; otherwise it returns a nonzero value.

## VBA Example



```
Sub ExtendFrameObject()

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

      ret = SapModel.FrameObj.AddByCoord(-180, 0, 180, -180, 0, 240, Name)



   'refresh window

      ret = SapModel.View.RefreshWindow



   'extend frame object

      ret = SapModel.EditFrame.Extend(Name, True, True, "7", "8")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[Trim](Trim.md)
