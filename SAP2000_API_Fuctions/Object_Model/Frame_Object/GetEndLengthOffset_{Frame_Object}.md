# GetEndLengthOffset

## Syntax

SapObject.SapModel.FrameObj.GetEndLengthOffset

## VB6 Procedure

Function GetEndLengthOffset(ByVal Name As String, ByRef AutoOffset As Boolean, ByRef Length1 As Double, ByRef Length2 As Double, ByRef rz As Double) As Long

## Parameters

Name

The name of an existing frame object.

AutoOffset

If this item is True, the end length offsets are automatically determined by the program from object connectivity.

Length1

The offset length along the 1-axis of the frame object at the I-End of the frame object. [L]

Length2

The offset along the 1-axis of the frame object at the J-End of the frame object. [L]

rz

The rigid zone factor.  This is the fraction of the end offset length assumed to be rigid for bending and shear deformations.

## Remarks

This function retrieves the frame object end offsets along the 1-axis of the object.

The function returns zero if the offsets are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameEndOffsets()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim AutoOffset As Boolean

      Dim Length1 As Double

      Dim Length2 As Double

      Dim rz As Double



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



   'get offsets

      ret = SapModel.FrameObj.GetEndLengthOffset("15", AutoOffset, Length1, Length2, rz)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetEndLengthOffset](SetEndLengthOffset.md)
