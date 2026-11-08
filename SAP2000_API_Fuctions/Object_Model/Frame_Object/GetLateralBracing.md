# GetLateralBracing

## Syntax

SapObject.SapModel.FrameObj.GetLateralBracing

## VB6 Procedure

Function GetLateralBracing(ByVal Name As String, ByRef NumberItems As Long, ByRef FrameName() As String, ByRef MyType() As Long, ByRef Loc() As Long, ByRef RD1() As Double, ByRef RD2() As Double, ByRef Dist1() As Double, ByRef Dist2() As Double) As Long

## Parameters

Name

The name of an existing frame object.

NumberItems

The total number of bracing assignments retrieved for the specified frame objects.

FrameName

This is an array that includes the name of the frame object associated with each bracing assignment.

MyType

This is an array that includes 1 or 2, indicating the bracing type assigned.

1 = Point bracing

2 = Uniform bracing

Loc

This is an array that includes 1, 2 or 3; indicating the bracing location.

1 = Top

2 = Bottom

3 = All (top and bottom)

RD1

This is an array that includes the relative location of the point bracing (when MyType = 1) or the relative location of the start of the uniform bracing (when MyType = 2).

RD2

This is an array that includes the relative location of the start of the uniform bracing (when MyType = 2).

This item does not apply for point bracing (when MyType = 1).

Dist1

This is an array that includes the actual location of the point bracing (when MyType = 1) or the actual location of the start of the uniform bracing (when MyType = 2). [L]

Dist2

This is an array that includes the actual location of the start of the uniform bracing (when MyType = 2). [L]

This item does not apply for point bracing (when MyType = 1).

## Remarks

This function retrieves the lateral bracing location assignments for frame objects.

The function returns zero if the assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameLateralBracing()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim FrameName() As String

      Dim MyType() As Long

      Dim Loc() As Long

      Dim RD1() As Double

      Dim RD2() As Double

      Dim Dist1() As Double

      Dim Dist2() As Double



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



   'assign frame lateral bracing

      ret = SapModel.FrameObj.SetLateralBracing("8", 1, 3, 0, 0)

      ret = SapModel.FrameObj.SetLateralBracing("8", 1, 3, 0.25, 0)

      ret = SapModel.FrameObj.SetLateralBracing("8", 1, 3, 1, 0)

      ret = SapModel.FrameObj.SetLateralBracing("8", 2, 1, 0.5, 1)



   'get frame lateral bracing

      ret = SapModel.FrameObj.GetLateralBracing("8", NumberItems, FrameName, MyType, Loc, RD1, RD2, Dist1, Dist2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetLateralBracing](SetLateralBracing.md)

[DeleteLateralBracing](DeleteLateralBracing.md)
