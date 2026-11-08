# Delete

## Syntax

SapObject.SapModel.PatternDef.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing joint pattern.

## Remarks

The function deletes the specified joint pattern. At least one joint pattern must be defined. The function will return an error if an attempt is made to delete the last joint pattern.

The function returns zero if the joint pattern is successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeletePattern()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'add joint pattern

      ret = SapModel.PatternDef.SetPattern("MyPattern")



   'delete joint pattern

      ret = SapModel.PatternDef.Delete("MyPattern")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetPattern](SetPattern.md)
