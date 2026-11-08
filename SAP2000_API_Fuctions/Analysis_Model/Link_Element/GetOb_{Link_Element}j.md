# GetObj

## Syntax

Sap2000.LinkElm.GetObj

## VB6 Procedure

Function GetObj(ByVal Name As String, ByRef Obj As String,
ByRef ObjType As Long) As Long

## Parameters

Name

The name of an existing link element.

Obj

The name of the object associated with the specified
link element. The type of object or item is determined from the ObjType
variable.

ObjType

A number indicating the type of object that is associated
with the point element.

**2**

Obj is a line object that
has a line spring assignment. The springs are modeled using link elements.

**3**

Obj is a area object that
has an area spring assignment. The springs are modeled using link elements.

**6**

Obj is a solid object
that has a surface spring assignment. The springs are modeled using link
elements.

**9**

Obj is a point object
that has a panel zone assignment. The specified link element is internally
added by the program at the point object (panel zone) location to model
the panel zone.

101

Obj is a point element

102

Obj is a line element

103

Obj is an area element

104

Obj is a plane element

106

Obj is a solid element

## Remarks

This function retrieves the object associated with a
specified link element.

The function returns zero if the object is successfully
retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetObjForLinkElm()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Obj As String

      Dim ObjType As Long

      Dim Vec() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'assign springs to frame

      ReDim Vec(2)

      ret = SapModel.FrameObj.SetSpring("8", 
 1, 1, 1, "", 1, 2, 0, Vec, 0, False, "Local")



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get object for link element

      ret = SapModel.LinkElm.GetObj("~1", 
 Obj, ObjType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Added point, line, area, and solid element as ObjType
options in v22.1.0.

Initial release in version 11.00.

## See Also
