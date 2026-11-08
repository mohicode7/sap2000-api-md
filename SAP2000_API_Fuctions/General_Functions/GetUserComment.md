# GetUserComment

## Syntax

SapObject.SapModel.GetUserComment

## VB6 Procedure

Function GetUserComment(ByRef Comment As String) As Long

## Parameters

Comment

The data in the user comments and log.

## Remarks

This function retrieves the data in the user comments and log.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetComments()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Comment As String



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



   'add comments

      ret = SapModel.SetUserComment("Testing the Sap2000 API.")

      ret = SapModel.SetUserComment("Adding a second comment.")

      ret = SapModel.SetUserComment("Adding a third comment.", 3)



   'get comments

      ret = SapModel.GetUserComment(Comment)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetUserComment](SetUserComment.md)
