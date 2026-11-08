# SetUserComment

## Syntax

SapObject.SapModel.SetUserComment

## VB6 Procedure

Function SetUserComment(ByVal Comment As String, Optional ByVal NumLines As Long = 1, Optional ByVal Replace As Boolean = False) As Long

## Parameters

Comment

The data to be added to the user comments and log.

NumLines

The number of carriage return and line feeds to be included before the specified comment. This item is ignored if Replace = True. It is also ignored if there are no existing comments.

Replace

If this item is True, all existing comments are replaced with the specified comment.

## Remarks

This function sets the user comments and log data.

The function returns zero if the data is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub AddCommentToLog()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetUserComment](GetUserComment.md)
