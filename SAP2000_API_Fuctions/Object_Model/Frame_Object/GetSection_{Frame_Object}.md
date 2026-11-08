# GetSection

## Syntax

SapObject.SapModel.FrameObj.GetSection

## VB6 Procedure

Function GetSection(ByVal Name As String, ByRef PropName
As String, ByRef SAuto As String) As Long

## Parameters

Name

The name of a defined frame object.

PropName

If no auto select list is assigned to the frame object,
this is the name of the frame section property assigned to the frame object.
If an auto select list is assigned to the frame object, this is the name
of the frame section property, within the auto select list, which is currently
being used as the analysis property for the frame object. If this item
is None, no frame section property is assigned to the frame object.

SAuto

This is the name of the auto select list assigned to
the frame object, if any. If this item is returned as a blank string,
no auto select list is assigned to the frame object.

## Remarks

This function retrieves the frame section property assigned
to a frame object.

The function returns zero if the frame object property
is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSectionProp()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropName As String

      Dim SAuto As String



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



   'get frame section property

      ret = SapModel.FrameObj.GetSection("3", 
 PropName, SAuto)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSectionNonPrismatic](GetSectionNonPrismatic.md)

[SetSection](SetSection.md)
