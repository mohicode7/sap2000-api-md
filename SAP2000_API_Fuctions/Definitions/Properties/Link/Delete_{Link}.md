# Delete

## Syntax

SapObject.SapModel.PropLink.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing link property.

## Remarks

The function deletes a specified link property.

The function returns zero if the link property is successfully deleted; otherwise it returns a nonzero value. It returns an error if the specified link property can not be deleted, for example, if it is being used by an existing link object.

## VBA Example



```
Sub DeleteLinkProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double

      Dim i As Long

      Dim NumberNames As Long

      Dim MyName() As String



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



   'add link properties

      ReDim DOF(5)

      ReDim Fixed(5)

      ReDim Ke(5)

      ReDim Ce(5)

      DOF(0) = True

      Ke(0) = 12

      ret = SapModel.PropLink.SetLinear("L1", DOF, Fixed, Ke, Ce, 0, 0)

      ret = SapModel.PropLink.SetLinear("L2", DOF, Fixed, Ke, Ce, 0, 0)



   'delete link property

      ret = SapModel.PropLink.Delete("L1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
