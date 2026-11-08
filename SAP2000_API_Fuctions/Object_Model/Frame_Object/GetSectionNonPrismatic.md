# GetSectionNonPrismatic

## Syntax

SapObject.SapModel.FrameObj.GetSectionNonPrismatic

## VB6 Procedure

Function GetSectionNonPrismatic(ByVal Name As String, ByRef PropName As String, ByRef sVarTotalLength As Double, ByRef sVarRelStartLoc As Double) As Long

## Parameters

Name

The name of a defined frame object.

PropName

The name of the nonprismatic frame section property assigned to the frame object.

sVarTotalLength

This is the total assumed length of the nonprismatic section. Enter 0 for this item to indicate that the section length is the same as the frame object length.

sVarRelStartLoc

This is the relative distance along the nonprismatic section to the I-End (start) of the frame object. This item is ignored when the sVarTotalLengthitem is 0.

## Remarks

This function retrieves the nonprismatic frame section property data assigned to a frame object.

The function returns zero if the nonprismatic frame object property data is successfully retrieved, otherwise it returns a nonzero value.

The function returns an error if the section property assigned to the frame object is not a nonprismatic property.

## VBA Example



```
Sub GetFrameSectionNonPrismatic()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim StartSec() As String

      Dim EndSec() As String

      Dim MyLength() As Double

      Dim MyType() As Long

      Dim EI33() As Long

      Dim EI22() As Long

      Dim Color As Long

      Dim Notes As String

      Dim GUID As String

      Dim PropName As String

      Dim sVarTotalLength As Double

      Dim sVarRelStartLoc As Double



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



   'set new I-type frame section property

      ret = SapModel.PropFrame.SetISection("ISEC1", "A992Fy50", 24, 8, 0.5, 0.3, 8, 0.5)



   'set new I-type frame section property

      ret = SapModel.PropFrame.SetISection("ISEC2", "A992Fy50", 20, 8, 0.5, 0.3, 8, 0.5)



   'set new nonprismatic frame section property

      ReDim StartSec(2)

      ReDim EndSec(2)

      ReDim MyLength(2)

      ReDim MyType(2)

      ReDim EI33(2)

      ReDim EI22(2)

      StartSec(0) = "ISEC2"

      EndSec(0) = "ISEC1"

      MyLength(0) = 60

      MyType(0) = 2

      EI33(0)= 2

      EI22(0)= 1



      StartSec(1) = "ISEC1"

      EndSec(1) = "ISEC1"

      MyLength(1) = 1

      MyType(1) = 1

      EI33(1)= 2

      EI22(1)= 1



      StartSec(2) = "ISEC1"

      EndSec(2) = "ISEC2"

      MyLength(2) = 60

      MyType(2) = 2

      EI33(2)= 2

      EI22(2)= 1



      ret = SapModel.PropFrame.SetNonPrismatic("NP1", 3, StartSec, EndSec, MyLength, MyType, EI33, EI22)



   'set frame section property

      ret = SapModel.FrameObj.SetSection("8", "NP1", Object, 0.1, 360)



   'get nonprismatic frame section property

      ret = SapModel.FrameObj.GetSectionNonPrismatic("8", PropName, sVarTotalLength, sVarRelStartLoc)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSection](GetSection_%7BFrame_Object%7D.md)

[SetSection](SetSection.md)
