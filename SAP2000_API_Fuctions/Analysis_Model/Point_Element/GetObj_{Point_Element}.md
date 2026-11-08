# GetObj

## Syntax

SapObject.SapModel.PointElm.GetObj

## VB6 Procedure

Function GetObj(ByVal Name As String, ByRef Obj As String,
ByRef ObjType As Long) As Long

## Parameters

Name

The name of an existing point element.

Obj

The name of the object or defined item associated with
the specified point element. The type of object or item is determined
from the ObjType variable.

ObjType

A number indicating the type of object or defined item
that is associated with the point element.

**1**

Obj is the point object
corresponding to the specified point element.

**2**

Obj is a line object that
is internally meshed by the program to create the specified point element.

**3**

Obj is an area object
that is internally meshed by the program to create the specified point
element.

**6**

Obj is a solid object
that is internally meshed by the program to create the specified point
element.

**9**

Obj is a point object
that has a panel zone assignment. The specified point element is internally
added by the program at the point object (panel zone) location to model
the panel zone. The specified point element does not directly correspond
to the point object returned; it is an added point at the same location
as the point object.

**21**

Obj is a defined diaphragm
constraint. The specified point element was internally added by the program
for application of auto wind and auto seismic loads.

107

Obj is a link element
that was generated as part of the analysis model.

## Remarks

This function retrieves the object or defined item associated
with a specified point element.

The function returns zero if the object is successfully
retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetObjForPointElm()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Obj As String

      Dim ObjType As Long



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
 3, 124, 3, 200)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'get object or defined item for point element "3"

      ret = SapModel.PointElm.GetObj("3", 
 Obj, ObjType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Added link element as an ObjType option in v22.1.0.

Initial release in version 11.00.

## See Also
