# GetInsertionPoint

## Syntax

SapObject.SapModel.LineElm.GetInsertionPoint

## VB6 Procedure

Function GetInsertionPoint(ByVal Name As String, ByRef Offset1() As Double, ByRef Offset2() As Double) As Long

## Parameters

Name

The name of an existing line element.

Offset1

This is an array of three joint offset distances, in the Global coordinate system, at the I-End of the line element. [L]

Offset1(0) = I-End offset in the global X-axis direction

Offset1(1) = I-End offset in the global Y-axis direction

Offset1(2) = I-End offset in the global Z-axis direction

Offset2

This is an array of three joint offset distances, in the Global coordinate system, at the J-End of the line element. [L]

Offset2(0) = J-End offset in the global X-axis direction

Offset2(1) = J-End offset in the global Y-axis direction

Offset2(2) = J-End offset in the global Z-axis direction

## Remarks

This function retrieves line element insertion point assignments. The assignments are reported as end joint offsets.

The function returns zero if the insertion point data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLineElmInsertionPoint()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim CardinalPoint As Long

      Dim Mirror2 As Boolean

      Dim StiffTransform As Boolean

      Dim Offset1() As Double

      Dim Offset2() As Double

      Dim CSys As String



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



   'assign frame object insertion point

      ReDim Offset1(2)

      ReDim Offset2(2)

      For i=0 To 2

         Offset1(i)=10 + i

         Offset2(i)=20 + i

      Next i

      ret = SapModel.FrameObj.SetInsertionPoint("15", 7, False, True, Offset1, Offset2)



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get line element insertion point

      ReDim Offset1(2)

      ReDim Offset2(2)

      ret = SapModel.LineElm.GetInsertionPoint("15-1", Offset1, Offset2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also
