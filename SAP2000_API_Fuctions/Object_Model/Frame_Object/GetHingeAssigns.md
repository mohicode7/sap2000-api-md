# GetHingeAssigns (Note:  Newer function available)

## Syntax

SapObject.SapModel.FrameObj.GetHingeAssigns

## VB6 Procedure

Function GetHingeAssigns(ByVal Name As String, ByRef
NumberHinges As Long, ByRef HingeNum() As Long, ByRef Prop() As String,
ByRef MyType() As Long, ByRef Behavior() As Long, ByRef Source() As String,
ByRef RD() As Double) As Long

## Parameters

Name

The name of an existing frame object.

NumberHinges

The number of hinge assignments on the specified frame
object.

HingeNum

An array that includes the hinge number for each hinge
on the frame object.

Prop

An array that includes the name of the generated hinge
property for each hinge on the frame object.

MyType

An array that specifies the type of hinge for each hinge
on the frame object. It is one of the following:

  1 = Axial
P

  2 = Shear
V2

  3 = Shear
V3

  4 = Torsion
T

  5 = Moment
M2

  6 = Moment
M3

  7 = Interacting
P-M2

  8 = Interacting
P-M3

  9 = Interacting
M2-M3

10 = Interacting P-M2-M3

11 = Fiber P-M2-M3

Behavior

An array that specifies the behavior of the hinge for
each hinge on the frame object. It is one of the following:

1 = Force controlled

2 = Deformation controlled

Source

An array that indicates the source of the generated
hinge property for each hinge on the frame object. The source is either
Auto or the name of a defined (not generated) hinge property.

RD

An array that indicates the relative distance of each
hinge along the frame object.

## Remarks

This function reports the hinge assignments for a specified
frame object.

The function returns zero if the assignment data is
successfully obtained; otherwise it returns a nonzero value.

## VBA Example

This example assumes that a file MyHinge.sdb exists.



```
Sub GetFrameHingeAssigns()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberHinges As Long

      Dim HingeNum() As Long

      Dim Prop() As String

      Dim MyType() As Long

      Dim Behavior() As Long

      Dim Source() As String

      Dim RD() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open an existing file

      FileName = "C:\SapAPI\MyHinge.sdb"

      ret = SapModel.File.OpenFile(FileName)



   'get hinge data for the frame object named 1

      ret = SapModel.FrameObj.GetHingeAssigns("1", 
 NumberHinges, HingeNum, Prop, MyType, Behavior, Source, RD)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.20.

## See Also

[GetHingeAssigns\_2](GetHingeAssigns_2.md)
