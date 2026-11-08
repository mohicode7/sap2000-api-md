# GetTypeRebar

## Syntax

SapObject.SapModel.PropFrame.GetTypeRebar

## VB6 Procedure

Function GetTypeRebar(ByVal Name As String, ByRef MyType As Long) As Long

## Parameters

Name

The name of an existing frame section property.

PropType

This is 0, 1 or 2, indicating the rebar design type.

0 = None

1 = Column

2 = Beam

## Remarks

This function retrieves the rebar design type for the specified frame section property.

The function returns zero if the type is successfully retrieved; otherwise it returns nonzero.

This function applies only to the following section property types. Calling this function for any other type of frame section property returns an error.

SECTION\_T = 3

SECTION\_ANGLE = 4

SECTION\_RECTANGULAR = 8

SECTION\_CIRCLE = 9

A nonzero rebar type is returned only if the frame section property has a concrete material.

## VBA Example



```
Sub GetFramePropRebarType()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long



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



   'set new frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'get rebar design type

      ret = SapModel.PropFrame.GetTypeRebar("R1", MyType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
