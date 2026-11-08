# ChangeName

## Syntax

SapObject.SapModel.PropLink.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined link property.

NewName

The new name for the link property.

## Remarks

This function changes the name of an existing link property.

The function returns zero if the new name is successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangeLinkPropName()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim Fixed() As Boolean

      Dim Ke() As Double

      Dim Ce() As Double



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



   'add link property

      ReDim DOF(5)

      ReDim Fixed(5)

      ReDim Ke(5)

      ReDim Ce(5)

      DOF(0) = True

      Ke(0) = 12

      ret = SapModel.PropLink.SetLinear("L1", DOF, Fixed, Ke, Ce, 0, 0)



   'change name of link property

      ret = SapModel.PropLink.ChangeName("L1", "MyLink")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also
