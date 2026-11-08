# Delete

## Syntax

SapObject.SapModel.PropFrame.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing frame section property.

## Remarks

The function deletes a specified frame section property.

The function returns zero if the frame section property is successfully deleted; otherwise it returns a nonzero value. It returns an error if the specified frame section property can not be deleted; for example, if it is being used by a frame object.

## VBA Example



```
Sub DeleteFrameProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim NumberNames As Long

      Dim MyName() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'set new frame section property

      ret = SapModel.PropFrame.SetGeneral("GEN1", "A992Fy50", 24, 14, 100, 80, 80, 4, 1000, 2400, 140, 200, 150, 220, 3, 5, -1, "API example", "Default")



   'get frame object names

      ret = SapModel.FrameObj.GetNameList(NumberNames, MyName)



   'set frame section property

      For i = 1 to NumberNames

         ret = SapModel.FrameObj.SetSection(MyName(i - 1), "GEN1")

      Next i



   'delete frame property

      ret = SapModel.PropFrame.Delete("FSEC1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
