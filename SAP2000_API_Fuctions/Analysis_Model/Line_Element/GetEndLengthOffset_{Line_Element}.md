# GetEndLengthOffset

## Syntax

SapObject.SapModel.LineElm.GetEndLengthOffset

## VB6 Procedure

Function GetEndLengthOffset(ByVal Name As String, ByRef Length1 As Double, ByRef Length2 As Double, ByRef rz As Double) As Long

## Parameters

Name

The name of an existing line element.

Length1

The offset length along the 1-axis of the line element at the I-End of the line element. [L]

Length2

The offset along the 1-axis of the line element at the J-End of the line element. [L]

rz

The rigid zone factor.  This is the fraction of the end offset length assumed to be rigid for bending and shear deformations.

## Remarks

This function retrieves the line element end offsets along the 1-axis of the element.

The function returns zero if the offsets are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmEndOffsets()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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



   'assign offsets to frame object

      ret = SapModel.FrameObj.SetEndLengthOffset("15", False, 12, 12, 0.5)



   'assign frame object auto mesh options

      ret = SapModel.FrameObj.SetAutoMesh("15", True, False, False, 2, 0)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get offsets for line element

      ret = SapModel.LineElm.GetEndLengthOffset("15-1", Length1, Length2, rz)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
