# GetNonPrismatic

## Syntax

SapObject.SapModel.PropFrame.GetNonPrismatic

## VB6 Procedure

Function GetNonPrismatic(ByVal Name As String, ByRef NumberItems As Long, ByRef StartSec() As String, ByRef EndSec() As String, ByRef MyLength() As Double, ByRef MyType() As Long, ByRef EI33() As Long, ByRef EI22() As Long, ByRef Color As Long, ByRef Notes As String, ByRef GUID As String) As Long

## Parameters

NumberItems

The number of segments assigned to the nonprismatic section.

StartSec

This is an array of the names of the frame section properties at the start of each segment.

EndSec

This is an array of the names of the frame section properties at the end of each segment.

MyLength

This is an array that includes the length of each segment. The length may be variable or absolute as indicated by the MyType item. [L] when length is absolute

MyType

This is an array of either 1 or 2, indicating the length type for each segment.

1 = Variable (relative length)

2 = Absolute

EI33, EI22

This is an array of either 1, 2 or 3, indicating the variation type for EI33 and EI22 in each segment.

1 = Linear

2 = Parabolic

3 = Cubic

Color

The display color assigned to the section.

Notes

The notes, if any, assigned to the section.

GUID

The GUID (global unique identifier), if any, assigned to the section.

## Remarks

This function retrieves frame section property data for nonprismatic (variable) sections.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePropNonPrismatic()

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



   'clear arrays

      ReDim StartSec(0)

      ReDim EndSec(0)

      ReDim MyLength(0)

      ReDim MyType(0)

      ReDim EI33(0)

      ReDim EI22(0)



   'get nonprismatic section data

      ret = SapModel.PropFrame.GetNonPrismatic("NP1", NumberItems, StartSec, EndSec, MyLength, MyType, EI33, EI22, Color, Notes, GUID)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetNonPrismatic](SetNonPrismatic.md)
