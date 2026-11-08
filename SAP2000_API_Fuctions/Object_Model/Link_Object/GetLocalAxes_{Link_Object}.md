# GetLocalAxes

## Syntax

SapObject.SapModel.LinkObj.GetLocalAxes

## VB6 Procedure

Function GetLocalAxes(ByVal Name As String, ByRef Ang As Double, ByRef Advanced As Boolean) As Long

## Parameters

Name

The name of an existing link object.

Ang

This is the angle that the local 2 and 3 axes are rotated about the positive local 1 axis, from the default orientation or, if the Advanced item is True, from the orientation determined by the plane reference vector. The rotation for a positive angle appears counter clockwise when the local +1 axis is pointing toward you. [deg]

Advanced

This item is True if the link object local axes orientation was obtained using advanced local axes parameters.

## Remarks

This function retrieves the local axis angle assignment for link objects.

The function returns zero if the assignment is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkLocalAxisAngle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim Ang As Double

      Dim Advanced As boolean



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



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name)



   'assign link local axis angle

      ret = SapModel.LinkObj.SetLocalAxes(Name, 30)



   'get link local axis angle

      ret = SapModel.LinkObj.GetLocalAxes(Name, Ang, Advanced)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetLocalAxes](SetLocalAxes_%7BLink_Object%7D.md)
